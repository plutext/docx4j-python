"""The engine's contract and the rules no Word-made fixture holds. CR-010 section 4.

The port of docx4j-core-tests' ``TrackedChangesTest`` (docx4j CR-043 section 12)
(docx4j CR-043 sections 12 and 13): the grouping, dates, a stale view, a piece
resolved on its own; the table, section and cell records (listed, accepted,
rejected, reported; records that change nothing), a section merged, cells
inserted and deleted, the legacy forms swept, and each filter (author and
type, a kind splitting a group, dates, a bookmark's extent, a control's
content); the unmeasured rules of CR-043 section 11 (U1, U2, U3, U15, a
hyperlink's relationship in a deleted row, the text box twin, an equation
structure) on handwritten markup. Then what this port adds: the typed islands
(CR-010 section 3.5), every story, and parts nobody reads left byte for byte.
"""

from __future__ import annotations

import datetime
import io
import zipfile

import pytest
from revision_helpers import (
    NS,
    body_xml,
    dele,
    fixture,
    ins,
    paragraphs,
    revisions,
    round_trip,
    run,
    text_of,
    unresolved,
    with_body,
)

from docx4j_py.model.content.errors import TrackedChangeError
from docx4j_py.model.revisions import Kind, TrackedChangeFilter
from docx4j_py.model.revisions.lister import Lister

# ---------------------------------------------------------------------------
# listing
# ---------------------------------------------------------------------------


def test_deletions_that_touch_are_one_change():
    """Check 22: a deleted paragraph is one change, a lone deleted mark one, a deletion across a mark one."""
    package = fixture("core-ts/check22/22-deleted.docx")
    changes = revisions.tracked_changes(package)
    assert all(change.type == "Deleted" for change in changes)
    assert [change.text for change in changes] == ["Two.\r", "\r", "six.\rSeven "]


def test_another_author_or_type_splits():
    """Check 23: another author splits a group; a deletion and an insertion are two changes."""
    package = with_body(
        "<w:p>" + dele(1, "A", "one ") + dele(2, "B", "two ") + ins(3, "B", "three ") + ins(4, "B", "four") + "</w:p>"
    )
    changes = revisions.tracked_changes(package)
    assert [change.text for change in changes] == ["one ", "two ", "three four"]
    assert len(changes[2].pieces) == 2
    assert changes[2].kind == "run_insertion"


def test_the_equations_list_as_word_counts_them():
    """CR-010 section 1: the equations fixture listed 0 changes; Word counts 12."""
    assert len(revisions.tracked_changes(fixture("tracked-changes-equations.docx"))) == 12


def test_a_text_box_is_listed_once_and_resolves_its_twin():
    """C22: listed from mc:Choice; accepting it resolves its mc:Fallback twin, paired by position."""
    package = fixture("cr043/cr043-tracked.docx")
    listed = [change for change in revisions.tracked_changes(package) if "C22 inserted" in change.text]
    assert len(listed) == 1
    inserted = listed[0]
    assert len(inserted.twins) == 1
    inserted.accept()
    xml = body_xml(package)
    first = xml.index("C22 inserted")
    assert xml.index("C22 inserted", first + 1) > 0, "both branches keep the text"
    for change in revisions.tracked_changes(package):
        assert "C22 inserted" not in change.text
    for change in Lister(package, False).list():
        assert "C22 inserted" not in change.text, "and the Fallback's insertion went with it"


# ---------------------------------------------------------------------------
# dates
# ---------------------------------------------------------------------------


def test_date_utc_is_preferred():
    package = with_body(
        '<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T10:00:00Z" '
        'w16du:dateUtc="2026-01-01T00:00:00Z">' + run("x") + "</w:ins></w:p>"
    )
    assert revisions.tracked_changes(package)[0].date == datetime.datetime(2026, 1, 1, tzinfo=datetime.UTC)


@pytest.mark.parametrize(
    ("place", "utc"),
    [
        ("Australia/Sydney", datetime.datetime(2025, 12, 31, 23, tzinfo=datetime.UTC)),  # AEDT, UTC+11
        ("America/New_York", datetime.datetime(2026, 1, 1, 15, tzinfo=datetime.UTC)),  # EST, UTC-5
        ("Asia/Kolkata", datetime.datetime(2026, 1, 1, 4, 30, tzinfo=datetime.UTC)),  # IST, UTC+5:30
    ],
)
def test_w_date_is_read_as_local_wall_clock_time(zone, place, utc):
    """docx4j ``TrackedChange.dateOf`` (CR-010 section 12.9): ``w:date``'s ``Z`` does not mean UTC.

    Without ``w16du:dateUtc``, ``w:date`` is Word's local wall-clock time, read
    in this process's zone; an explicit non-zero offset is read as it says. It
    was read as UTC until 12.9 (section 10.3 item 1).
    """
    zone(place)
    package = with_body('<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T10:00:00Z">' + run("x") + "</w:ins></w:p>")
    assert revisions.tracked_changes(package)[0].date == utc
    offset = with_body('<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T10:00:00+05:30">' + run("x") + "</w:ins></w:p>")
    assert revisions.tracked_changes(offset)[0].date == datetime.datetime(2026, 1, 1, 4, 30, tzinfo=datetime.UTC)


@pytest.mark.parametrize("place", ["Australia/Sydney", "America/New_York"])
def test_word_files_read_as_word_means_them(zone, place):
    """Word 365's ``w16du:dateUtc`` wins in any zone; Word 15's ``w:date`` alone is the local wall clock.

    The probe (Word 16, Jason's UTC+11): ``w:date="2026-10-16T10:17:00Z"``
    beside ``w16du:dateUtc="2026-10-15T23:17:00Z"``. Check 18's file, edited in
    Word 15, has no ``w16du:dateUtc``.
    """
    zone(place)
    probe = revisions.tracked_changes(fixture("cr043/cr043-tracked.docx"))[0]
    assert probe.date == datetime.datetime(2026, 10, 15, 23, 17, tzinfo=datetime.UTC)
    word15 = fixture("editor/revisions/revisions-word15.docx")
    first = revisions.tracked_changes(word15)[0]
    value = first.target.value()
    assert value.date_utc is None
    wall_clock = value.date.to_datetime().replace(tzinfo=None)
    assert first.date == wall_clock.astimezone(datetime.UTC)
    assert first.date != wall_clock.replace(tzinfo=datetime.UTC)


def test_a_filter_reads_dates_as_word_means_them(zone):
    """``TrackedChangeFilter``'s date range is over the dates as read: a Word 15 ``w:date`` on the local clock."""
    zone("Australia/Sydney")
    body = (
        '<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T10:00:00Z">' + run("old ") + "</w:ins>"
        + run("| ")
        + '<w:ins w:id="2" w:author="A" w:date="2026-01-02T10:00:00Z">' + run("new") + "</w:ins></w:p>"
    )
    package = with_body(body)
    # 10:00 AEDT on the 1st is 23:00 UTC on 31 December
    window = TrackedChangeFilter().between(
        datetime.datetime(2025, 12, 31, 22, tzinfo=datetime.UTC), datetime.datetime(2026, 1, 1, tzinfo=datetime.UTC)
    )
    revisions.reject_all(package, window)
    assert text_of(paragraphs(package)[0]) == "| new"


def test_no_date_is_none():
    package = with_body('<w:p><w:ins w:id="1" w:author="A">' + run("x") + "</w:ins></w:p>")
    assert revisions.tracked_changes(package)[0].date is None


# ---------------------------------------------------------------------------
# the contract
# ---------------------------------------------------------------------------


def test_a_view_is_stale_after_the_next_accept_or_reject():
    package = with_body("<w:p>" + ins(1, "A", "one") + run(" plain ") + ins(2, "A", "two") + "</w:p>")
    changes = revisions.tracked_changes(package)
    assert len(changes) == 2
    changes[1].accept()
    with pytest.raises(TrackedChangeError) as stale:
        changes[0].accept()
    assert stale.value.code == "tracking.gone"
    again = revisions.tracked_changes(package)
    assert len(again) == 1
    again[0].reject()
    assert text_of(paragraphs(package)[0]) == " plain two"


def test_a_content_api_view_is_stale_too():
    """CR-010 section 3.2: the content API's view raises ``tracking.gone`` as Phase F's did."""
    package = with_body("<w:p>" + ins(1, "A", "one") + run(" plain ") + ins(2, "A", "two") + "</w:p>")
    first, second = package.get_tracked_changes()
    second.accept()
    with pytest.raises(TrackedChangeError) as stale:
        first.reject()
    assert stale.value.code == "tracking.gone"
    assert "get_tracked_changes" in stale.value.hint


def test_a_piece_resolves_on_its_own():
    """A group's piece is a view of its own: rejected alone, it leaves the rest of the group pending."""
    package = with_body("<w:p>" + run("a ") + ins(1, "A", "one ") + ins(2, "A", "two") + "</w:p>")
    changes = revisions.tracked_changes(package)
    assert len(changes) == 1
    assert len(changes[0].pieces) == 2
    changes[0].pieces[0].reject()
    left = revisions.tracked_changes(package)
    assert [change.text for change in left] == ["two"]
    assert text_of(paragraphs(package)[0]) == "a two"


def test_a_filter_resolves_the_pieces_it_matches():
    """A plain callable is still a filter (CR-010 section 10.3 item 5); it applies to single changes (U19)."""
    package = with_body(
        '<w:p><w:pPr><w:rPr><w:ins w:id="2" w:author="A"/></w:rPr></w:pPr>'
        + run("kept ")
        + ins(1, "A", "inserted")
        + "</w:p><w:p>"
        + run("next")
        + "</w:p>"
    )
    assert len(revisions.tracked_changes(package)) == 1
    revisions.reject_all(package, lambda piece: piece.kind == "run_insertion")
    left = revisions.tracked_changes(package)
    assert [str(change.kind) for change in left] == ["paragraph_mark_insertion"]
    assert len(paragraphs(package)) == 2

    by_author = with_body("<w:p>" + ins(1, "A", "a1 ") + run("x ") + dele(2, "A", "a2 ") + run("y ") + ins(3, "B", "b1") + "</w:p>")
    assert revisions.accept_all(by_author, lambda piece: piece.author == "A") == 2
    assert [change.author for change in revisions.tracked_changes(by_author)] == ["B"]
    assert text_of(paragraphs(by_author)[0]) == "a1 x y b1"


# ---------------------------------------------------------------------------
# the table, section and cell records (docx4j CR-043 phase 2)
# ---------------------------------------------------------------------------

TABLE_RECORD = (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
    '<w:tblPrChange w:id="1" w:author="A"><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr></w:tblPrChange></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p>' + run("cell") + "</w:p></w:tc></w:tr></w:tbl>"
)


def test_a_tables_records_are_one_change_accepted_or_rejected():
    """docx4j ``tableRecords``: one change; accepted the properties stay, rejected they are restored; reported."""
    package = with_body(TABLE_RECORD + "<w:p>" + ins(2, "A", "text") + "</w:p>")
    changes = revisions.tracked_changes(package)
    assert [str(change.kind) for change in changes] == ["table_properties", "run_insertion"]
    assert changes[0].text == "cell\r\n"
    assert revisions.report(package).to_dict() == {
        "total": 2,
        "by_kind": {"run_insertion": 1, "table_properties": 1},
        "by_author": {"A": 2},
        "by_part": {"/word/document.xml": 2},
    }
    changes[0].reject()
    xml = body_xml(package)
    assert 'w:w="5000" w:type="pct"' in xml
    assert "tblPrChange" not in xml
    assert len(revisions.tracked_changes(package)) == 1

    accepted = with_body(TABLE_RECORD)
    assert revisions.accept_all(accepted) == 1
    xml = body_xml(accepted)
    assert 'w:w="0" w:type="auto"' in xml
    assert "tblPrChange" not in xml


def test_records_that_change_nothing_are_no_change():
    """docx4j ``recordsThatChangeNothing`` (checks 24, 26; U20): not listed, and swept either way."""
    package = with_body(
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblPrChange w:id="1" w:author="A"><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p>' + run("cell") + "</w:p></w:tc></w:tr></w:tbl><w:p/>"
    )
    assert revisions.tracked_changes(package) == []
    revisions.reject_all(package)
    assert "tblPrChange" not in body_xml(package)


def test_a_sections_record_rejected_is_merged():
    """docx4j ``sectionRecordMerges`` (C07, U13): the recorded members laid over it, a docGrid of pitch 0 meaning none."""
    package = with_body(
        "<w:p>" + run("text") + '</w:p><w:sectPr><w:headerReference w:type="default" r:id="rId99"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>'
        '<w:docGrid w:linePitch="299"/><w:sectPrChange w:id="1" w:author="A"><w:sectPr>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="2160" w:header="708" w:footer="708" w:gutter="0"/>'
        '<w:docGrid w:linePitch="0"/></w:sectPr></w:sectPrChange></w:sectPr>'
    )
    changes = revisions.tracked_changes(package)
    assert [str(change.kind) for change in changes] == ["section_properties"]
    assert changes[0].text == "text"
    assert package.outline().stats.tracked_changes == 1
    changes[0].reject()
    sect_pr = package.main_document_part.contents.body.sect_pr
    assert sect_pr.pg_mar.left == 2160
    assert sect_pr.pg_sz.w == 11906
    assert sect_pr.doc_grid is None
    assert sect_pr.sect_pr_change is None
    assert len(sect_pr.header_reference_or_footer_reference) == 1


CELLS = (
    '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>'
    '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p>' + run("one") + "</w:p></w:tc>"
    '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:cellIns w:id="1" w:author="A"/></w:tcPr><w:p>'
    + run("two")
    + "</w:p></w:tc></w:tr></w:tbl><w:p/>"
)


def test_u10_cells_inserted_and_deleted():
    """docx4j ``cellsInsertedAndDeleted`` (check 27): the cell removed widens the one before, the widths added."""
    package = with_body(CELLS)
    changes = revisions.tracked_changes(package)
    assert [str(change.kind) for change in changes] == ["cell_insertion"]
    changes[0].reject()
    xml = body_xml(package)
    assert "two" not in xml
    assert '<w:tcW w:w="3000" w:type="dxa"/><w:gridSpan w:val="2"/>' in xml

    accepted = with_body(CELLS)
    revisions.accept_all(accepted)
    xml = body_xml(accepted)
    assert "two" in xml
    assert "cellIns" not in xml

    deleted = with_body(CELLS.replace("cellIns", "cellDel"))
    revisions.accept_all(deleted)
    assert "two" not in body_xml(deleted)


def test_u8_u9_the_legacy_forms_are_swept():
    """docx4j ``legacyFormsSwept``: a cell merge, a numbering change and an inserted numPr, which nothing lists."""
    body = (
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:cellMerge w:id="1" w:author="A" w:vMerge="cont" w:vMergeOrig="rest"/></w:tcPr><w:p>'
        + run("merged")
        + "</w:p></w:tc></w:tr></w:tbl>"
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>'
        '<w:numberingChange w:id="2" w:author="A" w:original="%1."/></w:numPr></w:pPr>' + run("numbered") + "</w:p>"
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:ins w:id="3" w:author="A"/></w:numPr></w:pPr>'
        + run("listed")
        + "</w:p>"
    )
    rejected = with_body(body)
    assert revisions.tracked_changes(rejected) == []
    revisions.reject_all(rejected)
    xml = body_xml(rejected)
    assert "cellMerge" not in xml and "numberingChange" not in xml
    assert '<w:vMerge w:val="restart"/>' in xml
    assert xml.count("<w:numPr>") == 1

    accepted = with_body(body)
    revisions.accept_all(accepted)
    xml = body_xml(accepted)
    assert "cellMerge" not in xml and "numberingChange" not in xml and "<w:ins" not in xml
    assert "vMerge" not in xml
    assert xml.count("<w:numPr>") == 2


def test_a_fields_numbering_change_is_swept():
    """U9: a ``w:fldChar``'s ``w:numberingChange`` goes too, either way; the field stays."""
    field = (
        '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:numberingChange w:id="1" w:author="A" w:original="1."/>'
        '</w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    )
    for verb in (revisions.accept_all, revisions.reject_all):
        package = with_body(field)
        assert revisions.tracked_changes(package) == []
        verb(package)
        xml = body_xml(package)
        assert "numberingChange" not in xml
        assert xml.count("<w:fldChar ") == 2


def test_a_records_change_resolves_through_the_content_api():
    """Phase A's ``tracking.unsupported`` is gone: the content API's view accepts and rejects records, and counts them."""
    package = with_body(TABLE_RECORD + "<w:p>" + ins(2, "A", "text") + "</w:p>")
    first, _second = package.get_tracked_changes()
    assert first.kind == "table_properties"
    first.reject()
    assert "tblPrChange" not in body_xml(package)
    assert [change.kind for change in package.get_tracked_changes()] == ["run_insertion"]

    counted = with_body(TABLE_RECORD + "<w:p>" + ins(2, "A", "text") + "</w:p>")
    assert counted.body.accept_all() == 2
    assert counted.body.get_tracked_changes() == []


# ---------------------------------------------------------------------------
# filters (docx4j CR-043 phase 2) and the report
# ---------------------------------------------------------------------------


def test_filter_by_author_and_type():
    """docx4j ``filterByAuthorAndType``."""
    package = with_body("<w:p>" + ins(1, "A", "a1 ") + run("x ") + dele(2, "A", "a2 ") + run("y ") + ins(3, "B", "b1") + "</w:p>")
    assert revisions.accept_all(package, TrackedChangeFilter().authors("A")) == 2
    assert text_of(paragraphs(package)[0])[:7] == "a1 x y "
    left = revisions.tracked_changes(package)
    assert [change.author for change in left] == ["B"]

    typed = with_body("<w:p>" + ins(1, "A", "a1 ") + run("x ") + dele(2, "A", "a2 ") + "</w:p>")
    revisions.reject_all(typed, TrackedChangeFilter(types=["Added"]))
    assert [change.type for change in revisions.tracked_changes(typed)] == ["Deleted"]


def test_filter_by_kind_splits_a_group():
    """docx4j ``filterByKindSplitsAGroup`` (U19): only the group's matching pieces resolve."""
    package = with_body(
        '<w:p><w:pPr><w:rPr><w:ins w:id="2" w:author="A"/></w:rPr></w:pPr>'
        + run("kept ")
        + ins(1, "A", "inserted")
        + "</w:p><w:p>"
        + run("next")
        + "</w:p>"
    )
    assert len(revisions.tracked_changes(package)) == 1
    revisions.reject_all(package, TrackedChangeFilter().kinds(Kind.RUN_INSERTION))
    left = revisions.tracked_changes(package)
    assert [str(change.kind) for change in left] == ["paragraph_mark_insertion"]
    assert len(paragraphs(package)) == 2


def test_filter_by_date():
    """docx4j ``filterByDate``: at or after the start, before the end; a naive datetime is UTC."""
    body = (
        '<w:p><w:ins w:id="1" w:author="A" w16du:dateUtc="2026-01-01T00:00:00Z">' + run("old ") + "</w:ins>"
        + run("| ")
        + '<w:ins w:id="2" w:author="A" w16du:dateUtc="2026-06-01T00:00:00Z">' + run("new") + "</w:ins></w:p>"
    )
    package = with_body(body)
    revisions.reject_all(package, TrackedChangeFilter().between(datetime.datetime(2026, 3, 1, tzinfo=datetime.UTC), None))
    assert text_of(paragraphs(package)[0]) == "old | "
    assert len(revisions.tracked_changes(package)) == 1

    naive = with_body(body)
    revisions.reject_all(naive, TrackedChangeFilter(date_to=datetime.datetime(2026, 3, 1)))  # noqa: DTZ001 - naive is UTC
    assert text_of(paragraphs(naive)[0]) == "| new"
    undated = with_body("<w:p>" + ins(1, "A", "no date") + "</w:p>")
    since = datetime.datetime(2000, 1, 1)  # noqa: DTZ001 - naive is read as UTC
    assert revisions.accept_all(undated, TrackedChangeFilter(date_from=since)) == 0


def test_filter_within_a_bookmark():
    """docx4j ``filterWithinABookmark``: the changes between its start and its end."""
    package = with_body(
        "<w:p>" + ins(1, "A", "before ") + run("| ") + '<w:bookmarkStart w:id="9" w:name="here"/>' + ins(2, "A", "inside")
        + '<w:bookmarkEnd w:id="9"/>' + run(" | ") + ins(3, "A", "after") + "</w:p>"
    )
    assert len(revisions.tracked_changes(package)) == 3
    assert [change.text for change in revisions.tracked_changes(package, TrackedChangeFilter().within_bookmark("here"))] == [
        "inside"
    ]
    revisions.reject_all(package, TrackedChangeFilter().within_bookmark("here"))
    assert text_of(paragraphs(package)[0]) == "before |  | after"
    with pytest.raises(ValueError, match="no bookmark named 'nowhere'"):
        revisions.accept_all(package, TrackedChangeFilter(within_bookmark="nowhere"))


def test_filter_within_a_content_control():
    """docx4j ``filterWithinAControl``: the model's control, or the content API's view of it."""
    from docx4j_py.wml import SdtBlock

    body = (
        "<w:p>" + run("one ") + "</w:p><w:sdt><w:sdtPr/><w:sdtContent><w:p>" + ins(1, "A", "inside")
        + "</w:p></w:sdtContent></w:sdt><w:p>" + ins(2, "A", "outside") + "</w:p>"
    )
    package = with_body(body)
    sdt = next(item for item in package.main_document_part.contents.body.content if isinstance(item, SdtBlock))
    revisions.reject_all(package, TrackedChangeFilter().within(sdt))
    assert [change.text for change in revisions.tracked_changes(package)] == ["outside"]

    viewed = with_body(body)
    control = viewed.body.content_controls[0]
    assert revisions.accept_all(viewed, TrackedChangeFilter(within=control)) == 1
    assert [change.text for change in revisions.tracked_changes(viewed)] == ["outside"]


def test_filter_by_part():
    """A part, or its name: the probe's endnotes resolved, every other story left."""
    package = fixture("cr043/cr043-tracked.docx")
    endnotes = package.main_document_part.endnotes_part
    before = revisions.report(package).by_part
    assert revisions.accept_all(package, TrackedChangeFilter(parts=["/word/endnotes.xml"])) == before["/word/endnotes.xml"]
    after = revisions.report(package).by_part
    assert "/word/endnotes.xml" not in after
    assert after == {name: count for name, count in before.items() if name != "/word/endnotes.xml"}
    assert revisions.reject_all(package, TrackedChangeFilter().parts(package.main_document_part)) == before["/word/document.xml"]
    assert set(revisions.report(package).by_part) == {"/word/header2.xml", "/word/footer2.xml", "/word/header5.xml", "/word/footer5.xml"}
    assert endnotes is package.main_document_part.endnotes_part


def test_a_filter_with_no_condition_takes_everything_but_sweeps_nothing():
    """A filter is a filter even with no condition: every change resolved, nothing unlisted swept (docx4j's rule)."""
    package = with_body(TABLE_RECORD.replace('w:w="5000" w:type="pct"', 'w:w="0" w:type="auto"') + "<w:p>" + ins(2, "A", "text") + "</w:p>")
    assert revisions.accept_all(package, TrackedChangeFilter()) == 1
    assert "tblPrChange" in body_xml(package), "records that change nothing are swept only with no filter"
    assert repr(TrackedChangeFilter(authors=["A"]).kinds("run_insertion")) == (
        "TrackedChangeFilter(authors=['A'], kinds=['run_insertion'])"
    )


def test_the_report_counts_by_kind_author_and_part():
    """docx4j ``TrackedChanges.report``: the probe, every story; ``to_dict`` plain for JSON."""
    import json

    package = fixture("cr043/cr043-tracked.docx")
    report = revisions.report(package)
    assert report.total == len(revisions.tracked_changes(package)) == 37
    data = report.to_dict()
    assert json.loads(json.dumps(data)) == data
    assert list(data) == ["total", "by_kind", "by_author", "by_part"]
    assert data["by_author"] == {"Author A": 35, "CR-043 base": 2}
    assert data["by_part"] == {
        "/word/document.xml": 28,
        "/word/endnotes.xml": 5,
        "/word/header2.xml": 1,
        "/word/footer2.xml": 1,
        "/word/header5.xml": 1,
        "/word/footer5.xml": 1,
    }
    assert data["by_kind"]["table_properties"] == 1 and data["by_kind"]["section_properties"] == 1
    assert list(data["by_kind"]) == [str(kind) for kind in revisions.Kind if str(kind) in data["by_kind"]]
    revisions.accept_all(package)
    assert revisions.report(package).to_dict() == {"total": 0, "by_kind": {}, "by_author": {}, "by_part": {}}


# ---------------------------------------------------------------------------
# the unmeasured rules (docx4j CR-043 section 11)
# ---------------------------------------------------------------------------

DELETED_MARK = '<w:pPr><w:rPr><w:del w:id="{0}" w:author="A"/></w:rPr></w:pPr>'


def test_u1_no_join_across_a_container_boundary():
    """U1, revised by docx4j's text-left probe (``e55a83b8e``, CR-010 section 13.3a): docx4j ``noJoinAcrossAContainerBoundary``.

    A deleted mark with a table next joins its paragraph's content into the
    table's first cell's first paragraph; a cell's last mark and a story's
    last mark are dropped, with no join.
    """
    package = with_body(
        "<w:p>" + DELETED_MARK.format(1) + run("before a table") + "</w:p>"
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p>'
        + DELETED_MARK.format(2)
        + run("a cell's last")
        + "</w:p></w:tc></w:tr></w:tbl>"
        "<w:p>" + DELETED_MARK.format(3) + run("the story's last") + "</w:p><w:sectPr/>"
    )
    assert revisions.accept_all(package) == 3
    assert len(paragraphs(package)) == 1, '"before a table" went into the cell'
    assert revisions.tracked_changes(package) == []
    assert "w:del " not in body_xml(package)
    body = package.main_document_part.contents.body
    assert body.sect_pr is not None
    cell = body.content[0].content[0].content[0]
    assert len(cell.content) == 1
    assert text_of(cell.content[0]) == "before a tablea cell's last"


def test_a_join_keeps_the_next_paragraph():
    """A deleted mark joins: the result keeps the next paragraph's properties and paraId (C10b); markers between go into it."""
    package = with_body(
        '<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:del w:id="1" w:author="A"/></w:rPr></w:pPr>'
        + run("first ")
        + "</w:p>"
        '<w:bookmarkEnd w:id="9"/>'
        '<w:p w14:paraId="00000002"><w:pPr><w:jc w:val="right"/></w:pPr>' + run("second") + "</w:p>"
    )
    revisions.accept_all(package)
    ps = paragraphs(package)
    assert len(ps) == 1
    assert ps[0].para_id == "00000002"
    assert ps[0].p_pr.p_style is None
    assert text_of(ps[0]) == "first second"
    assert len(package.main_document_part.contents.body.content) == 1
    assert any(type(item).__name__ == "CTMarkupRange" for item in ps[0].content), "the marker went into the paragraph"


# ---------------------------------------------------------------------------
# check 54's three Accept All rules: a mark taken away with no paragraph next
# (CR-010 section 13; docx4j 29dbf5964, its CR-043 section 17)
# ---------------------------------------------------------------------------


def marked(revision: str, identifier: int, properties: str = "") -> str:
    """A ``w:pPr`` holding `properties`, its mark inserted (``ins``) or deleted (``del``)."""
    return f'<w:pPr>{properties}<w:rPr><w:{revision} w:id="{identifier}" w:author="A"/></w:rPr></w:pPr>'


CENTRED = '<w:jc w:val="center"/>'
HEADING = '<w:pStyle w:val="Heading1"/>'
LISTED = '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
A_TABLE = '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p>' + run("cell") + "</w:p></w:tc></w:tr></w:tbl>"
AFTER = "<w:p>" + run("after") + "</w:p>"


def control(*blocks: str) -> str:
    return '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>' + "".join(blocks) + "</w:sdtContent></w:sdt>"


def alternate(*blocks: str) -> str:
    """A kept block-level ``mc:AlternateContent``, `blocks` in the Choice Word takes and a paragraph in its Fallback."""
    return (
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="w14">' + "".join(blocks) + "</mc:Choice>"
        "<mc:Fallback><w:p>" + run("fallback") + "</w:p></mc:Fallback></mc:AlternateContent>"
    )


def whole(accept: bool, properties: str, mark_id: int = 1) -> str:
    """A paragraph its resolution takes away whole: its text and mark deleted (to accept) or inserted (to reject)."""
    if accept:
        return "<w:p>" + marked("del", mark_id, properties) + dele(2, "A", "gone") + "</w:p>"
    return "<w:p>" + marked("ins", mark_id, properties) + ins(2, "A", "new") + "</w:p>"


def tail(accept: bool, properties: str) -> str:
    """A paragraph its resolution takes the mark and the end of the text away from, ``HEAD `` left."""
    if accept:
        return "<w:p>" + marked("del", 1, properties) + run("HEAD ") + dele(2, "A", "gone") + "</w:p>"
    return "<w:p>" + marked("ins", 1, properties) + run("HEAD ") + ins(2, "A", "new") + "</w:p>"


def resolved(body: str, accept: bool):
    package = with_body(body)
    (revisions.accept_all if accept else revisions.reject_all)(package)
    assert unresolved(package) == []
    return package


def assert_view_agrees(body: str, accept: bool) -> None:
    """The text model's view (CR-011, ``view.py``) reads as the copy the resolver resolved, in text and in Markdown."""
    from docx4j_py.model.text import extract, extract_resolved, markdown, plain_text

    view = "accepted" if accept else "original"
    for render in (plain_text, markdown):
        walk = render(extract(with_body(body), view=view))
        copy = render(extract_resolved(resolved(body, accept), view=view))
        assert walk == copy, f"{render.__name__}, {view}: walk {walk!r}, copy {copy!r}"


def blocks_of(package) -> list:
    return list(package.main_document_part.contents.body.content)


@pytest.mark.parametrize("accept", [True, False], ids=["deletion-accepted", "insertion-rejected"])
@pytest.mark.parametrize("following", ["table", "control"])
def test_an_emptied_paragraph_with_a_table_or_a_control_next_goes(following, accept):
    """Check 54 case 9 and the text-left probe's case 3: the paragraph and its mark taken away, a table or control next.

    The paragraph goes, its ``w:jc`` with it, and what is next follows what
    came before. Both ways, as docx4j's ``joinWithNext`` takes the mark away
    whichever resolution took it; Word's Reject All of both probes reached
    none of these.
    """
    block = A_TABLE if following == "table" else control("<w:p>" + run("BODY") + "</w:p>")
    body = "<w:p>" + run("before") + "</w:p>" + whole(accept, CENTRED) + block + AFTER
    package = resolved(body, accept)
    assert [type(item).__name__ for item in blocks_of(package)] == ["P", {"table": "Tbl"}.get(following, "SdtBlock"), "P"]
    assert "center" not in body_xml(package)
    assert_view_agrees(body, accept)


@pytest.mark.parametrize("accept", [True, False], ids=["deletion-accepted", "insertion-rejected"])
def test_text_left_before_a_table_joins_the_first_cells_first_paragraph(accept):
    """The text-left probe's case 1: the content goes to the front of the first cell's first paragraph, which keeps its properties.

    The paragraph goes, and a range marker between it and the table goes
    into the cell's paragraph with its content, as a join carries one.
    """
    table = (
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>'
        '<w:tc><w:p><w:pPr><w:jc w:val="right"/></w:pPr>' + run("cell") + "</w:p><w:p>" + run("second") + "</w:p></w:tc>"
        "<w:tc><w:p>" + run("next cell") + "</w:p></w:tc></w:tr></w:tbl>"
    )
    body = tail(accept, HEADING) + '<w:bookmarkEnd w:id="9"/>' + table + AFTER
    package = resolved(body, accept)
    items = blocks_of(package)
    assert [type(item).__name__ for item in items] == ["Tbl", "P"]
    first, second = items[0].content[0].content[0].content
    assert text_of(first) == "HEAD cell"
    assert first.p_pr.p_style is None and first.p_pr.jc.val.value == "right"
    assert text_of(second) == "second"
    assert any(type(item).__name__ == "CTMarkupRange" for item in first.content), "the marker went with the content"
    assert_view_agrees(body, accept)


def test_paragraphs_joining_one_another_then_a_table_join_its_first_cell_in_order():
    """Two marks taken away before a table, text left in each: both paragraphs' text, in order, then the cell's own."""
    first = "<w:p>" + marked("del", 1) + run("one ") + "</w:p>"
    body = first + tail(True, HEADING).replace('w:id="1"', 'w:id="3"', 1) + A_TABLE + AFTER
    package = resolved(body, True)
    items = blocks_of(package)
    assert [type(item).__name__ for item in items] == ["Tbl", "P"]
    assert text_of(items[0].content[0].content[0].content[0]) == "one HEAD cell"
    assert_view_agrees(body, True)


def test_a_paragraph_left_holding_range_markers_only_is_emptied():
    """docx4j's ``hasTextOrContent``: range markers are not content, so the paragraph goes, its empty bookmark with it.

    The view reads it so too, where docx4j's, asking ``remainsIn``, keeps the
    paragraph and reads an empty line the resolved copy does not have (CR-010
    section 14.4).
    """
    body = (
        "<w:p>" + marked("del", 1, CENTRED) + '<w:bookmarkStart w:id="5" w:name="_GoBack"/><w:bookmarkEnd w:id="5"/>'
        + dele(2, "A", "gone") + "</w:p>" + A_TABLE + AFTER
    )
    package = resolved(body, True)
    assert [type(item).__name__ for item in blocks_of(package)] == ["Tbl", "P"]
    assert_view_agrees(body, True)


@pytest.mark.parametrize("accept", [True, False], ids=["deletion-accepted", "insertion-rejected"])
@pytest.mark.parametrize("emptied", [True, False], ids=["emptied", "text-left"])
def test_a_controls_last_paragraph_keeps_a_bare_mark_text_left_or_not(emptied, accept):
    """Check 54 case 13 and the text-left probe's case 2: the last paragraph of a control's content, its mark taken away.

    Its properties go, its style and number among them, its text left with
    it: a bare mark (``e55a83b8e`` widened ``29dbf5964``'s rule to text left).
    """
    body = control("<w:p>" + run("first") + "</w:p>", (whole if emptied else tail)(accept, LISTED)) + AFTER
    package = resolved(body, accept)
    sdt = blocks_of(package)[0]
    first, last = sdt.sdt_content.content
    assert text_of(first) == "first"
    assert last.p_pr is None
    assert text_of(last) == ("" if emptied else "HEAD ")
    assert_view_agrees(body, accept)


@pytest.mark.parametrize("accept", [True, False], ids=["deletion-accepted", "insertion-rejected"])
def test_text_left_before_a_control_takes_its_first_paragraphs_properties(accept):
    """Check 54 case 19 and the text-left probe's case 5: nothing joins; the paragraph in the control's first paragraph's properties, a copy."""
    body = tail(accept, HEADING) + control(
        "<w:p><w:pPr>" + CENTRED + "</w:pPr>" + run("BODY") + "</w:p>", "<w:p>" + run("second") + "</w:p>"
    )
    package = resolved(body, accept)
    head, sdt = blocks_of(package)
    first = sdt.sdt_content.content[0]
    assert text_of(head) == "HEAD "
    assert head.p_pr.p_style is None and head.p_pr.jc.val.value == "center" and head.p_pr.r_pr is None
    assert head.p_pr is not first.p_pr and head.p_pr.parent is head
    assert text_of(first) == "BODY" and first.p_pr.jc.val.value == "center"
    assert_view_agrees(body, accept)


def test_a_custom_xml_element_is_a_control_and_a_table_first_gives_nothing():
    """docx4j's ``isHolder`` takes a block ``w:customXml`` as a control; a control or cell whose first block is a table gives nothing.

    Text left before a control whose content starts with a table, or before a
    table whose first cell does, keeps its own properties, the mark only
    dropped.
    """
    nested = (
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>' + A_TABLE + "<w:p>" + run("after nested")
        + "</w:p></w:tc></w:tr></w:tbl>"
    )
    tabled = control(A_TABLE, "<w:p><w:pPr>" + CENTRED + "</w:pPr>" + run("BODY") + "</w:p>")
    custom = '<w:customXml w:element="e"><w:p><w:pPr>' + CENTRED + "</w:pPr>" + run("X") + "</w:p></w:customXml>"
    body = tail(True, HEADING) + tabled + tail(True, HEADING) + custom + tail(True, HEADING) + nested + AFTER
    package = resolved(body, True)
    first, _sdt, second, _xml, third, _tbl, _after = blocks_of(package)
    assert first.p_pr.p_style.val == "Heading1" and first.p_pr.jc is None, "past a table: the mark only dropped"
    assert second.p_pr.p_style is None and second.p_pr.jc.val.value == "center"
    assert text_of(third) == "HEAD " and third.p_pr.p_style.val == "Heading1"
    assert_view_agrees(body, True)


@pytest.mark.parametrize("accept", [True, False], ids=["deletion-accepted", "insertion-rejected"])
def test_a_cells_a_stories_and_a_custom_xml_elements_last_paragraphs_keep_their_marks_dropped(accept):
    """U1 as it stands: a cell's, a story's and a custom XML element's last paragraph keep their properties, the mark only dropped.

    The text-left probe's case 4, a cell's last mark deleted, made Word
    collapse a one-cell table: not followed, as docx4j does not follow it.
    """
    cell = (
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p>' + run("first") + "</w:p>"
        + whole(accept, LISTED) + "</w:tc></w:tr></w:tbl>"
    )
    custom = '<w:customXml w:element="e"><w:p>' + run("first") + "</w:p>" + whole(accept, LISTED) + "</w:customXml>"
    body = cell + custom + tail(accept, LISTED)
    package = resolved(body, accept)
    tbl, xml, last = blocks_of(package)
    for paragraph in (tbl.content[0].content[0].content[-1], xml.content[-1]):
        assert text_of(paragraph) == "" and paragraph.p_pr.p_style.val == "ListParagraph"
    assert text_of(last) == "HEAD " and last.p_pr.p_style.val == "ListParagraph"
    assert_view_agrees(body, accept)


@pytest.mark.parametrize("accept", [True, False], ids=["deletion-accepted", "insertion-rejected"])
def test_a_kept_block_level_branch_next_reads_as_words_branch(accept):
    """CR-010 section 13.2: a kept ``mc:AlternateContent`` next is read as the first block of the branch Word takes.

    A table or a control there takes an emptied paragraph away, and a control
    gives one with text left its first paragraph's properties. No content
    joins into a kept branch: text left before a table there, or a paragraph
    there, and the mark is only dropped (docx4j reads no further than the
    element).
    """
    in_control = control("<w:p><w:pPr>" + CENTRED + "</w:pPr>" + run("BODY") + "</w:p>")
    body = (
        whole(accept, CENTRED) + alternate(A_TABLE)
        + whole(accept, CENTRED) + alternate(in_control)
        + tail(accept, HEADING) + alternate(in_control)
        + tail(accept, HEADING) + alternate(A_TABLE)
        + whole(accept, CENTRED) + alternate("<w:p>" + run("branch") + "</w:p>")
        + AFTER
    )
    package = resolved(body, accept)
    items = blocks_of(package)
    names = [type(item).__name__ for item in items]
    assert names == ["AlternateContent", "AlternateContent", "P", "AlternateContent", "P", "AlternateContent", "P",
                     "AlternateContent", "P"]
    assert text_of(items[2]) == "HEAD " and items[2].p_pr.jc.val.value == "center" and items[2].p_pr.p_style is None
    assert text_of(items[4]) == "HEAD " and items[4].p_pr.p_style.val == "Heading1"
    assert text_of(items[6]) == "" and items[6].p_pr.jc.val.value == "center"
    assert_view_agrees(body, accept)


def _plan(body: str, accepted: bool):
    """The text model's plan (``ResolvedView``) over a body, with the body's blocks."""
    from docx4j_py.model.text.view import ResolvedView

    package = with_body(body)
    return ResolvedView(Lister(package, False).list(), accepted), blocks_of(package)


#: Check 54 case 15's markup, as Word saved it: a deleted mark before a control deleted whole in Word's
#: ``w:customXmlDelRange*`` form, its paragraph deleted whole inside it.
DELETED_CONTROL = (
    "<w:p>" + marked("del", 1, HEADING) + run("HEAD ") + dele(2, "A", "tail") + "</w:p>"
    '<w:customXmlDelRangeStart w:id="3" w:author="A"/>'
    '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:customXmlDelRangeEnd w:id="3"/>'
    "<w:p>" + marked("del", 4) + dele(5, "A", "inside") + "</w:p>"
    '<w:customXmlDelRangeStart w:id="6" w:author="A"/></w:sdtContent></w:sdt>'
    '<w:customXmlDelRangeEnd w:id="6"/>'
    "<w:p><w:pPr>" + CENTRED + "</w:pPr>" + run("BODY") + "</w:p>"
)


def test_the_view_reads_a_control_deleted_whole_as_absent_accepted_and_present_rejected():
    """CR-010 section 13.1: ``CONTENT_CONTROL_DELETION`` (CR-013 ``e7f18a7``) in the text model's view, docx4j's ``ResolvedView``.

    Accepted, the control is gone with what it holds, and the mark before it
    joins the paragraph after (check 54 cases 15 and 17); rejected, the control
    stays, its content in it, not unwrapped as an inserted one rejected is.
    """
    accepted, blocks = _plan(DELETED_CONTROL, True)
    sdt = next(item for item in blocks if type(item).__name__ == "SdtBlock")
    assert accepted.removes(sdt) and not accepted.unwraps(sdt)
    original, blocks = _plan(DELETED_CONTROL, False)
    sdt = next(item for item in blocks if type(item).__name__ == "SdtBlock")
    assert not original.removes(sdt) and not original.unwraps(sdt)
    for accept in (True, False):
        assert_view_agrees(DELETED_CONTROL, accept)
    package = resolved(DELETED_CONTROL, True)
    assert [text_of(p) for p in blocks_of(package)] == ["HEAD BODY"]


def test_a_filter_takes_a_control_deleted_whole_by_its_kind():
    """``TrackedChangeFilter`` over ``Kind.CONTENT_CONTROL_DELETION``, as core-ts's filter takes it since its ``222cc20``.

    The filter here takes any kind by name, so the kind CR-013 ``e7f18a7``
    added needed no entry; this holds it: only the control goes, the
    deletions around it still listed.
    """
    package = with_body(DELETED_CONTROL)
    taken = revisions.accept_all(package, TrackedChangeFilter(kinds=[Kind.CONTENT_CONTROL_DELETION]))
    assert taken == 1
    assert [type(item).__name__ for item in blocks_of(package)] == ["P", "P"]
    assert {str(change.kind) for change in revisions.tracked_changes(package)} == {"deletion"}


def test_the_view_removes_a_table_whose_every_row_it_removes():
    """Check 54 case 6, a gap docx4j found in its view: the rows deleted, the table goes, and the heading joins the paragraph after it."""
    rows = "".join(
        f'<w:tr><w:trPr><w:del w:id="{10 + k}" w:author="A"/></w:trPr><w:tc><w:p>'
        + marked("del", 20 + k) + dele(30 + k, "A", f"cell {k}") + "</w:p></w:tc></w:tr>"
        for k in range(2)
    )
    body = (
        "<w:p>" + marked("del", 1, HEADING) + run("HEAD ") + dele(2, "A", "tail") + "</w:p>"
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' + rows + "</w:tbl>" + AFTER
    )
    accepted, blocks = _plan(body, True)
    assert accepted.removes(blocks[1])
    original, blocks = _plan(body, False)
    assert not original.removes(blocks[1])
    for accept in (True, False):
        assert_view_agrees(body, accept)
    assert [text_of(p) for p in blocks_of(resolved(body, True))] == ["HEAD after"]


def test_the_view_decides_a_long_run_of_emptied_paragraphs_before_a_table_without_recursing():
    """Rule 1 over 2,000 paragraphs deleted whole, one after another, before a table: each goes, in the view as in the copy."""
    body = "".join(whole(True, CENTRED, mark_id=k) for k in range(100, 2100)) + A_TABLE + AFTER
    plan, blocks = _plan(body, True)
    assert all(plan.removes(block) for block in blocks[:2000])
    assert not plan.removes(blocks[2000])
    assert [type(item).__name__ for item in blocks_of(resolved(body, True))] == ["Tbl", "P"]


def test_u2_a_field_straddling_an_edge_resolves_as_runs():
    xml = (
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:del w:id="1" w:author="A"><w:r><w:delText>1</w:delText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "<w:r><w:delText> after</w:delText></w:r></w:del></w:p>"
    )
    package = with_body(xml)
    changes = revisions.tracked_changes(package)
    assert [str(change.kind) for change in changes] == ["run_deletion"]
    rejected = with_body(xml)
    revisions.reject_all(rejected)
    assert 'w:fldCharType="end"' in body_xml(rejected)
    revisions.accept_all(package)
    assert 'w:fldCharType="end"' not in body_xml(package)


def test_c01_an_inserted_field_rejected_goes_whole():
    """C01: Word's form for an inserted field, rejected: begin, instruction, separator, result and end all go.

    CR-003 section 22.3 found Phase F leaving a lone ``separate`` and its result.
    """
    xml = (
        "<w:p>" + run("before ")
        + '<w:ins w:id="1" w:author="A"><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        "<w:r><w:instrText> PAGE </w:instrText></w:r></w:ins>"
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>'
        '<w:ins w:id="2" w:author="A"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:ins></w:p>'
    )
    package = with_body(xml)
    changes = revisions.tracked_changes(package)
    assert [(str(c.kind), c.text) for c in changes] == [("field_insertion", "1")]
    revisions.reject_all(package)
    assert "fldChar" not in body_xml(package) and "PAGE" not in body_xml(package)
    assert text_of(paragraphs(package)[0]) == "before "
    accepted = with_body(xml)
    revisions.accept_all(accepted)
    xml_after = body_xml(accepted)
    assert xml_after.count("<w:fldChar ") == 3 and "<w:ins " not in xml_after


def test_c11_bookmarks_touching_removed_content_go():
    """A bookmark with an end in removed content goes whole (C11); one around removed content goes too (C11b, U15)."""
    xml = (
        "<w:p>" + run("a ") + '<w:bookmarkStart w:id="10" w:name="half"/>' + run("b ")
        + '<w:del w:id="1" w:author="A"><w:r><w:delText>c </w:delText></w:r><w:bookmarkEnd w:id="10"/></w:del>'
        + run("d ")
        + '<w:bookmarkStart w:id="11" w:name="around"/>' + ins(2, "A", "e") + '<w:bookmarkEnd w:id="11"/>'
        + '<w:bookmarkStart w:id="12" w:name="point"/><w:bookmarkEnd w:id="12"/></w:p>'
    )
    package = with_body(xml)
    revisions.accept_all(package)
    out = body_xml(package)
    assert "half" not in out
    assert "around" in out
    assert "point" in out
    rejected = with_body(xml)
    revisions.reject_all(rejected)
    out = body_xml(rejected)
    assert "half" in out
    assert "around" not in out and 'w:id="11"' not in out
    assert "point" in out


def test_u4_a_hyperlinks_relationship_goes_with_its_deleted_row():
    from docx4j_py.openpackaging.parts.namespaces import Namespaces

    package = with_body("<w:p/>")
    rels = package.main_document_part.get_relationships_part(True)
    rel = rels.add_external_relationship(Namespaces.HYPERLINK, "https://www.docx4java.org/")
    package.main_document_part.set_xml(
        f"<w:document {NS}><w:body><w:p>" + run("kept") + "</w:p>"
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>'
        "<w:tr><w:tc><w:p>" + run("row one") + "</w:p></w:tc></w:tr>"
        '<w:tr><w:trPr><w:del w:id="1" w:author="A"/></w:trPr><w:tc><w:p>'
        f'<w:hyperlink r:id="{rel.id}"><w:del w:id="2" w:author="A"><w:r><w:delText>link</w:delText></w:r></w:del>'
        "</w:hyperlink></w:p></w:tc></w:tr></w:tbl><w:p/></w:body></w:document>"
    )
    assert rels.get_relationship_by_id(rel.id) is not None
    revisions.accept_all(package)
    assert "row one" in body_xml(package)
    assert rels.get_relationship_by_id(rel.id) is None


def test_u3_a_removed_comment_takes_its_replies_and_entries():
    """A removed comment's entries in commentsExtended, commentsIds and commentsExtensible go, and its reply; then the parts."""
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "content"))
    from docx4j_py.openpackaging.content_types import ContentTypes
    from docx4j_py.openpackaging.parts.default_xml_part import DefaultXmlPart
    from docx4j_py.openpackaging.parts.namespaces import Namespaces
    from docx4j_py.openpackaging.parts.wml import CommentsExtendedPart, CommentsPart

    package = with_body(
        "<w:p>" + run("before ")
        + '<w:ins w:id="1" w:author="A"><w:commentRangeStart w:id="20"/><w:commentRangeStart w:id="21"/>'
        "<w:r><w:t>commented</w:t></w:r></w:ins>"
        '<w:commentRangeEnd w:id="20"/><w:commentRangeEnd w:id="21"/>'
        '<w:r><w:commentReference w:id="20"/></w:r><w:r><w:commentReference w:id="21"/></w:r>'
        + run(" after")
        + "</w:p>"
    )
    main = package.main_document_part
    w14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    comments = CommentsPart()
    comments.set_xml(
        f"<w:comments {w} {w14}>"
        '<w:comment w:id="20" w:author="A"><w:p w14:paraId="0000000A">' + run("the comment") + "</w:p></w:comment>"
        '<w:comment w:id="21" w:author="B"><w:p w14:paraId="0000000B">' + run("the reply") + "</w:p></w:comment>"
        "</w:comments>"
    )
    main.add_target_part(comments)
    w15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
    extended = CommentsExtendedPart()
    extended.set_xml(
        f"<w15:commentsEx {w15}>"
        '<w15:commentEx w15:paraId="0000000A" w15:done="0"/>'
        '<w15:commentEx w15:paraId="0000000B" w15:paraIdParent="0000000A" w15:done="0"/></w15:commentsEx>'
    )
    main.add_target_part(extended)
    cid = 'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"'
    cex = 'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"'
    for name, content_type, relationship, xml in (
        (
            "/word/commentsIds.xml",
            ContentTypes.WORDPROCESSINGML_COMMENTS_IDS,
            Namespaces.COMMENTS_IDS,
            (
                f"<w16cid:commentsIds {cid}>"
                '<w16cid:commentId w16cid:paraId="0000000A" w16cid:durableId="1000000A"/>'
                '<w16cid:commentId w16cid:paraId="0000000B" w16cid:durableId="1000000B"/></w16cid:commentsIds>'
            ),
        ),
        (
            "/word/commentsExtensible.xml",
            ContentTypes.WORDPROCESSINGML_COMMENTS_EXTENSIBLE,
            Namespaces.COMMENTS_EXTENSIBLE,
            (
                f"<w16cex:commentsExtensible {cex}>"
                '<w16cex:commentExtensible w16cex:durableId="1000000A"/>'
                '<w16cex:commentExtensible w16cex:durableId="1000000B"/></w16cex:commentsExtensible>'
            ),
        ),
    ):
        part = DefaultXmlPart(name, content_type, relationship)
        part.set_xml(xml)
        main.add_target_part(part)
    ids, extensible = main.comments_ids_part, main.comments_extensible_part

    revisions.reject_all(package)
    assert text_of(paragraphs(package)[0]) == "before  after"
    assert "comment" not in body_xml(package)
    assert comments.contents.comment == []
    assert extended.contents.comment_ex == []
    assert "commentId " not in ids.get_xml() and "1000000B" not in ids.get_xml()
    assert "1000000A" not in extensible.get_xml()
    # the last comment gone, its parts go with it (CR-010 section 12.8, fix 2)
    for part in (comments, extended, ids, extensible):
        assert package.get_part(part.part_name) is None, part.part_name
    assert main.comments_part is None and main.comments_ids_part is None


def _commented(body: str, comments: dict[int, str]) -> object:
    """A package whose body is `body`, with a comment entry (and its people.xml person) per id."""
    from docx4j_py.openpackaging.parts.wml import CommentsPart

    package = with_body(body)
    part = CommentsPart()
    part.set_xml(
        f"<w:comments {NS}>"
        + "".join(
            f'<w:comment w:id="{cid}" w:author="A"><w:p w14:paraId="0000000{cid % 10}">{run(text)}</w:p></w:comment>'
            for cid, text in comments.items()
        )
        + "</w:comments>"
    )
    package.main_document_part.add_target_part(part)
    return package


def _comment_ids(package) -> list[int]:
    part = package.main_document_part.comments_part
    return [] if part is None else [c.id for c in part.contents.comment]


def test_a_comment_whose_markers_lie_either_side_of_a_rejected_insertion_goes_whole():
    """Fix 2 (CR-010 section 12.8; docx4j ``d18693760``): Word 365 removes it, markers, reference and entry.

    Word 2010 and Word 15 kept it, losing only its range markers, which was the
    rule here before. A comment whose range still holds text stays, and so does
    one whose range was empty to begin with.
    """
    package = _commented(
        "<w:p>" + run("before ")
        + '<w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>'
        + '<w:commentRangeStart w:id="2"/>' + ins(10, "A", "inserted") + '<w:commentRangeEnd w:id="2"/>'
        + '<w:r><w:commentReference w:id="2"/></w:r>'
        + '<w:commentRangeStart w:id="3"/>' + run("kept ") + ins(11, "A", "gone") + '<w:commentRangeEnd w:id="3"/>'
        + '<w:r><w:commentReference w:id="3"/></w:r>'
        + run(" after") + "</w:p>",
        {1: "empty before", 2: "on the insertion", 3: "on kept text"},
    )
    revisions.reject_all(package)
    xml = body_xml(package)
    assert 'w:id="2"' not in xml
    assert _comment_ids(package) == [1, 3]
    assert xml.count('w:commentRangeStart w:id="1"') == 1 and xml.count('w:commentRangeStart w:id="3"') == 1
    assert _comment_ids(round_trip(package)) == [1, 3]


def test_an_accepted_deletion_takes_a_comment_whose_markers_lie_either_side_of_it():
    """U21, by symmetry with the rejected insertion: unmeasured in Word, docx4j's rule."""
    package = _commented(
        "<w:p>" + run("before ")
        + '<w:commentRangeStart w:id="2"/>' + dele(10, "A", "deleted") + '<w:commentRangeEnd w:id="2"/>'
        + '<w:r><w:commentReference w:id="2"/></w:r>' + run(" after") + "</w:p>",
        {2: "on the deletion"},
    )
    revisions.accept_all(package)
    assert "comment" not in body_xml(package)
    assert text_of(paragraphs(package)[0]) == "before  after"


def test_with_the_last_comment_its_parts_go_and_people_stays_while_a_revision_is_left():
    """docx4j ``Removals.dropEmptyCommentParts``: people.xml goes only when no tracked change is left."""
    package = fixture("python/artefact8/8-tracked-changes.docx")
    main = package.main_document_part
    parts = [main.comments_part, main.comments_extended_part, main.comments_ids_part]
    people = main.people_part
    second = next(change for change in revisions.tracked_changes(package) if change.text == "second")
    second.reject()
    assert len(revisions.tracked_changes(package)) == 8
    for part in parts:
        assert package.get_part(part.part_name) is None, part.part_name
        assert package.content_type_manager.get_override_content_type(part.part_name) is None
        assert main.relationships_part.get_rel(part.part_name) is None
    assert main.comments_part is None and main.comments_extended_part is None and main.comments_ids_part is None
    assert main.people_part is people and package.get_part(people.part_name) is people
    saved = round_trip(package)
    assert saved.main_document_part.comments_part is None
    assert saved.main_document_part.people_part is not None
    assert "comment" not in body_xml(saved)


def test_a_dry_run_removes_no_comment_part_and_the_real_reject_does():
    """CR-010 section 10.3 item 3: a dry run's body resolves locally, and touches no part outside its own."""
    package = fixture("python/artefact8/8-tracked-changes.docx")
    before = {name: package.get_part(name) for name in ("/word/comments.xml", "/word/people.xml")}
    with package.dry_run() as trial:
        assert trial.body.reject_all() > 0
        assert trial.body.get_tracked_changes() == []
    main = package.main_document_part
    for name, part in before.items():
        assert package.get_part(name) is part, name
    assert main.comments_part is before["/word/comments.xml"] and len(main.comments_part.contents.comment) == 1
    assert len(revisions.tracked_changes(package)) == 9

    revisions.reject_all(package)
    for name in before:
        assert package.get_part(name) is None, name


# -- a comment's range across paragraphs (CR-043 U23 and U24; docx4j a8bc6cda9) --


def _spanning(
    cid: int,
    *,
    marks: tuple[bool, bool] = (True, True),
    plain: str = "",
    start_inside: bool = False,
    deleted: bool = False,
    authors: tuple[str, str] = ("A", "A"),
) -> str:
    """U23's shape: two paragraphs whose text is inserted (or deleted), commented whole, and a paragraph after.

    The range starts in the first paragraph (before its ``w:ins``, or inside
    it) and ends in the second, after its ``w:ins`` and after `plain` text, if
    any; `marks` says which of the two paragraph marks carry the revision too.
    """
    change, kind = (dele, "del") if deleted else (ins, "ins")
    start = f'<w:commentRangeStart w:id="{cid}"/>'

    def mark(on: bool, rid: int, author: str) -> str:
        return f'<w:pPr><w:rPr><w:{kind} w:id="{rid}" w:author="{author}"/></w:rPr></w:pPr>' if on else ""

    first = change(100 + cid, authors[0], f"first {cid}")
    if start_inside:
        first = first.replace("<w:r>", start + "<w:r>", 1)
    return (
        f"<w:p>{mark(marks[0], 300 + cid, authors[0])}{'' if start_inside else start}{first}</w:p>"
        f"<w:p>{mark(marks[1], 400 + cid, authors[1])}{change(200 + cid, authors[1], f'second {cid}')}"
        + (run(plain) if plain else "")
        + f'<w:commentRangeEnd w:id="{cid}"/><w:r><w:commentReference w:id="{cid}"/></w:r></w:p>'
        f"<w:p>{run(f'after {cid}')}</w:p>"
    )


_CONTROL = (
    '<w:p><w:commentRangeStart w:id="2"/>' + run("control") + '<w:commentRangeEnd w:id="2"/>'
    '<w:r><w:commentReference w:id="2"/></w:r></w:p>'
)


def _markers(package, cid: int) -> list[int]:
    """How many range starts, range ends and references of comment `cid` the body holds."""
    xml = body_xml(package)
    return [xml.count(f'<w:{name} w:id="{cid}"/>') for name in ("commentRangeStart", "commentRangeEnd", "commentReference")]


@pytest.mark.parametrize("start_inside", [False, True], ids=["start-before", "start-inside"])
@pytest.mark.parametrize("deleted", [False, True], ids=["insertions-rejected", "deletions-accepted"])
def test_a_comment_across_paragraphs_goes_whole_once_the_joins_leave_its_range_empty(deleted, start_inside):
    """U23 (Word 365's Reject All of two inserted paragraphs commented whole; docx4j ``a8bc6cda9``).

    The comment's start and end are in different paragraphs, so it is judged at
    the flush, after the marks taken away have joined them into one paragraph
    with nothing with text between: it goes whole, markers, reference and
    entry, and the control comment stays. Accepting two deleted paragraphs is
    docx4j's rule by symmetry (unmeasured); a start inside the first
    revision, the content API's own form, is deferred the same way.
    """
    package = _commented(
        "<w:p>" + run("before") + "</w:p>" + _spanning(1, deleted=deleted, start_inside=start_inside) + _CONTROL,
        {1: "on both paragraphs", 2: "the control"},
    )
    (revisions.accept_all if deleted else revisions.reject_all)(package)
    assert [text_of(p) for p in paragraphs(package)] == ["before", "after 1", "control"]
    assert _markers(package, 1) == [0, 0, 0]
    assert _markers(package, 2) == [1, 1, 1]
    assert _comment_ids(package) == [2]
    assert _comment_ids(round_trip(package)) == [2]


def test_a_comment_across_paragraphs_stays_while_a_mark_or_text_is_left_in_its_range():
    """U24 (docx4j's reading, unmeasured): a range still crossing a paragraph mark is kept, as one holding text.

    Every paragraph's inserted text is rejected. Comment 1's marks are inserted
    too, so it goes (U23). Comment 3's marks are not, comment 4's first mark is
    not (the one between its start and end), and comment 5's range keeps text
    of its own: all three stay, each with its markers and reference.
    """
    package = _commented(
        "<w:p>" + run("before") + "</w:p>"
        + _spanning(1)
        + _spanning(3, marks=(False, False))
        + _spanning(4, marks=(False, True))
        + _spanning(5, plain="kept")
        + _CONTROL,
        {1: "goes", 2: "the control", 3: "no mark taken", 4: "a mark left between", 5: "text left"},
    )
    revisions.reject_all(package)
    assert _comment_ids(package) == [2, 3, 4, 5]
    assert _markers(package, 1) == [0, 0, 0]
    for cid in (3, 4, 5):
        assert _markers(package, cid) == [1, 1, 1], cid
    assert "keptafter 5" in [text_of(p) for p in paragraphs(package)]
    assert _comment_ids(round_trip(package)) == [2, 3, 4, 5]


def test_a_comment_across_paragraphs_in_a_vml_text_box_is_judged_by_its_text():
    """A comment whose range crosses paragraphs in a VML text box is judged at the flush by the text left in it.

    Comment 1's range is left empty and goes; comment 3's still holds the
    text "kept" and stays. Until CR-012 Phase C the box was a wildcard tree,
    resolved as a typed island and written back before the flush, which read
    its text by element name (CR-010 section 3.5); it is typed now, and the
    outcome is the same.
    """
    v = 'xmlns:v="urn:schemas-microsoft-com:vml"'
    box = (
        '<w:p><w:r><w:pict><v:shape id="box" style="width:200pt;height:100pt"><v:textbox><w:txbxContent>'
        + "<w:p>" + run("in the box") + "</w:p>" + _spanning(1) + _spanning(3, plain="kept")
        + "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"
    )
    package = _commented(box + _CONTROL, {1: "goes", 2: "the control", 3: "text left"})
    main = package.main_document_part
    main.set_xml(main.get_xml().replace("<w:document ", f"<w:document {v} ", 1))
    revisions.reject_all(package)
    assert unresolved(package) == []
    assert _comment_ids(package) == [2, 3]
    assert _markers(package, 1) == [0, 0, 0]
    assert _markers(package, 3) == [1, 1, 1]


def test_a_dry_run_judges_no_comment_across_paragraphs_and_the_real_reject_does():
    """CR-010 section 10.3 item 3: a dry run's body resolves locally; no comment entry or part is touched."""
    package = fixture("cr043/comments/c43-u23-reject-all-two-paragraph-insertion-commented.docx")
    main = package.main_document_part
    parts = [main.comments_part, main.comments_extended_part, main.comments_ids_part, main.people_part]
    xml = {part.part_name: part.get_xml() for part in parts}
    assert main.contents is not None  # read, so that both sides below are this package's marshalling
    body = body_xml(package)
    with package.dry_run() as trial:
        assert trial.body.reject_all() == 1
        assert trial.body.get_tracked_changes() == []
    assert body_xml(package) == body
    for part in parts:
        assert package.get_part(part.part_name) is part, part.part_name
        assert part.get_xml() == xml[part.part_name], part.part_name
    assert _comment_ids(package) == [0, 5]

    revisions.reject_all(package)
    assert _comment_ids(package) == [5]
    assert main.comments_part is parts[0] and main.people_part is parts[3]


@pytest.mark.parametrize("route", ["engine", "content"])
def test_one_change_at_a_time_judges_the_comment_across_paragraphs_at_each_flush(route):
    """The single-change path: ``TrackedChange.reject()`` flushes, and the deferred judgement runs there too.

    Each paragraph by another author, so two changes. The second rejected first
    joins its paragraph with the next; the first's text is still in the range,
    so the comment stays. The first rejected then empties it, and it goes.
    """
    package = _commented(_spanning(1, authors=("A", "B")) + _CONTROL, {1: "on both paragraphs", 2: "the control"})

    def listed():
        return revisions.tracked_changes(package) if route == "engine" else package.body.get_tracked_changes()

    assert [change.text for change in listed()] == ["first 1\r", "second 1\r"]
    listed()[1].reject()
    assert _comment_ids(package) == [1, 2]
    assert _markers(package, 1) == [1, 1, 1]
    remaining = listed()
    assert [change.text for change in remaining] == ["first 1\r"]
    remaining[0].reject()
    assert listed() == []
    assert _comment_ids(package) == [2]
    assert _markers(package, 1) == [0, 0, 0]
    assert [text_of(p) for p in paragraphs(package)] == ["after 1", "control"]


def test_math_structures():
    """An inserted equation structure: accepted, its control keeps the recorded run properties; rejected, it goes whole."""
    package = fixture("tracked-changes-equations.docx")
    inserted = revisions.tracked_changes(package)[1]
    assert inserted.text.endswith("s-1.")
    assert any(piece.kind == "math_structure_insertion" for piece in inserted.pieces)
    inserted.accept()
    xml = body_xml(package)
    s_sup = xml.index("<m:sSup>")
    ctrl_pr = xml[xml.index("<m:ctrlPr>", s_sup) : xml.index("</m:ctrlPr>", s_sup)]
    assert "w:ins" not in ctrl_pr
    assert "Cambria Math" in ctrl_pr

    rejected = fixture("tracked-changes-equations.docx")
    revisions.tracked_changes(rejected)[1].reject()
    xml = body_xml(rejected)
    first_equation = xml[xml.index("<m:oMathPara>") : xml.index("</m:oMathPara>")]
    assert "<m:sSup>" not in first_equation, "the inserted structure went whole"


# ---------------------------------------------------------------------------
# what this port adds: typed islands, every story, parts left alone
# ---------------------------------------------------------------------------


def _alternate_content_xml(package: object) -> str:
    from docx4j_py.child import is_alternate_content
    from docx4j_py.fragments import to_xml
    from docx4j_py.traversal import iter_nodes
    from docx4j_py.wml import R

    body = package.main_document_part.contents.body
    for node in iter_nodes(body, mce="all"):
        if isinstance(node, R):
            for item in node.content:
                if is_alternate_content(item):
                    return to_xml(item)
    raise AssertionError("no mc:AlternateContent")


def test_a_typed_island_leaves_the_rest_of_its_branch_unchanged():
    """CR-010 section 3.5 and risk 4: only the w:txbxContent is rewritten; everything around it is as it was."""
    import re

    package = fixture("cr043/cr043-tracked.docx")
    before = _alternate_content_xml(package)
    revisions.accept_all(package)
    after = _alternate_content_xml(package)
    boxes = re.compile(r"<w:txbxContent>.*?</w:txbxContent>", re.DOTALL)
    # the fragment's namespace declarations are what its content uses: w16du
    # goes with the revisions that carried w16du:dateUtc, and nothing else does
    declarations = re.compile(r'\sxmlns:\w+="[^"]*"')
    assert declarations.sub("", boxes.sub("BOX", before)) == declarations.sub("", boxes.sub("BOX", after)), (
        "the shape, the VML and both branches untouched"
    )
    inside = boxes.findall(after)
    assert len(inside) == 2
    for box in inside:
        assert "<w:ins " not in box and "<w:del " not in box
        assert "C22 inserted." in box and "DELETEME" not in box


def test_a_text_box_with_no_revision_is_never_made_an_island():
    """A document whose text boxes nobody edited keeps them exactly: no island, no rewrite."""
    from docx4j_py.model.revisions import islands

    package = fixture("cr043/cr043-accept-all.docx")  # Word's save: no revision left anywhere
    made: list = []
    original = islands.Island.__init__

    def counting(self, node, holder):
        made.append(node)
        original(self, node, holder)

    islands.Island.__init__ = counting
    try:
        assert revisions.tracked_changes(package) == []
    finally:
        islands.Island.__init__ = original
    assert made == []


def test_every_story_is_resolved_by_the_package_verbs():
    """CR-010 section 3.3: headers, footers, notes and comments with the body; Word's probe log counts none left."""
    for verb in (revisions.accept_all, revisions.reject_all):
        package = fixture("cr043/cr043-tracked.docx")
        verb(package)
        assert revisions.tracked_changes(package) == []
        saved = round_trip(package)
        for part in saved.main_document_part.header_parts() + saved.main_document_part.footer_parts():
            assert "<w:ins " not in part.get_xml() and "<w:del " not in part.get_xml()
        endnotes = saved.main_document_part.endnotes_part.get_xml()
        assert "<w:ins " not in endnotes and "<w:del " not in endnotes and "rPrChange" not in endnotes


def test_a_bodys_verbs_resolve_its_own_story_and_leave_the_others():
    """``Body.accept_all()`` is the body's story, with Word's side effects; the headers keep theirs."""
    package = fixture("cr043/cr043-tracked.docx")
    package.body.accept_all()
    left = {str(change.part.part_name) for change in package.get_tracked_changes()}
    # C06's deleted section break, accepted, took the first section's headers with it (D7)
    assert left == {"/word/header5.xml", "/word/footer5.xml", "/word/endnotes.xml"}
    assert package.get_part("/word/header2.xml") is None


def test_a_note_goes_with_its_reference():
    """C12: the endnote whose reference was deleted is removed from its part when the deletion is accepted."""
    package = fixture("cr043/cr043-tracked.docx")
    endnotes = package.main_document_part.endnotes_part.contents.endnote
    before = len(endnotes)
    revisions.accept_all(package)
    assert len(package.main_document_part.endnotes_part.contents.endnote) == before - 1


def test_a_picture_part_goes_with_its_last_reference():
    """C04, C05 and D7: Word's saves keep no unreferenced image; accepted, the deleted picture's part goes."""

    def images(package: object) -> int:
        return sum(1 for part in package.parts.values() if (part.content_type or "").startswith("image/"))

    package = fixture("cr043/cr043-tracked.docx")
    before = images(package)
    revisions.accept_all(package)
    after = images(package)
    assert after == before - 1
    with zipfile.ZipFile(io.BytesIO(package.save())) as archive:
        types = archive.read("[Content_Types].xml").decode("utf-8")
        names = set(archive.namelist())
    for part in package.parts.values():
        assert part.part_name.store_name in names or not part.part_name.store_name
    assert types.count("image/") <= types.count("Extension=") + after


def test_listing_does_not_unmarshal_a_part_with_no_revisions():
    """A story part whose bytes hold no revision markup is not read, so it is still written back byte for byte."""
    package = fixture("core-ts/check22/22-deleted.docx")
    main = package.main_document_part
    others = [p for p in (main.footnotes_part, main.endnotes_part, main.comments_part, *main.header_parts()) if p]
    revisions.tracked_changes(package)
    assert all(not part.is_unmarshalled for part in others)


def test_nothing_resolvable_is_left_and_the_result_saves():
    package = fixture("cr043/cr043-tracked.docx")
    revisions.reject_all(package)
    assert unresolved(package) == []
    round_trip(package)
