"""docx4j CR-067's date-forms probe, replayed here. CR-025 Phase B2.

docx4j's ``DateFormsProbeWordTest`` at ``c58e36e64``: the probe
(``tests/fixtures/date-forms-cr067/``, copied from docx4j) holds three stored
forms, a date alone (``2024-02-03``), a time with no offset
(``2025-06-07T08:09:10``) and a UTC time (``2025-01-02T03:04:05Z``), each in a
core property (D01 to D06), a revision's ``w:date`` (R01 to R03) and a date
control's ``w:fullDate`` (B01 to B03). Word 365 in Sydney's zone read them, on
F9, in its balloons and in the picker, as ``results.txt`` records, and its save
rewrote them (``date-forms-r1-word365.docx``):

* **the fields** (item 1, docx4j's D13): a core property's bare value is UTC, so
  ``CREATEDATE`` and ``SAVEDATE`` give Word's results through the text model's
  ``fields: update``; ``PRINTDATE`` keeps its cached result there (CR-025 Q6),
  as docx4j's text model keeps it;
* **the revisions** (item 2): the local clock in every form, a date alone local
  midnight, the ``Z`` ignored;
* **the date controls**: the nominal day in every form;
* **the save** (item 3, the user's "item 3 as recommended"): a date alone and a
  time with no offset written as Word's save wrote them, the nominal clock to
  the minute labelled ``Z``; the ``Z`` form left as read, where Word's save cut
  its seconds (the user's decision).
"""

from __future__ import annotations

import datetime
import hashlib
import io
import re
import sys
import zipfile
import zoneinfo
from pathlib import Path

import pytest
from lxml import etree

from docx4j_py import load
from docx4j_py.model.fields import FieldContext, evaluate, parse_w3cdtf
from docx4j_py.model.revisions import tracked_changes
from docx4j_py.model.text import extract, plain_text

ROOT = Path(__file__).resolve().parents[2]
PROBE = ROOT / "tests" / "fixtures" / "date-forms-cr067"
SYDNEY = "Australia/Sydney"
UTC = datetime.UTC
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

pytestmark = pytest.mark.filterwarnings("ignore::docx4j_xsdata.exceptions.ConverterWarning")

#: R01 to R03 as Word showed them in Sydney (``results.txt`` step 3), as instants: 3/02/2024 12:00:00 AM
#: (UTC+11), 7/06/2025 8:09 AM (UTC+10), 2/01/2025 3:04 AM (UTC+11, the ``Z`` ignored). docx4j's
#: ``theRevisionsAreWordsLocalClock``.
WORD_REVISIONS = [
    datetime.datetime(2024, 2, 2, 13, 0, tzinfo=UTC),
    datetime.datetime(2025, 6, 6, 22, 9, 10, tzinfo=UTC),
    datetime.datetime(2025, 1, 1, 16, 4, 5, tzinfo=UTC),
]

#: B01 to B03: "in each case the picker highlighted the corresponding date" (``results.txt`` step 4).
WORD_PICKER_DAYS = [datetime.date(2024, 2, 3), datetime.date(2025, 6, 7), datetime.date(2025, 1, 2)]

#: What Word's save wrote for every ``w:date`` and ``w:fullDate``, in document order.
WORD_SAVE_FORMS = ["2024-02-03T00:00:00Z", "2025-06-07T08:09:00Z", "2025-01-02T03:04:00Z"]

#: What a save here writes: Word's for the two forms Word does not write, the ``Z`` form as read.
OUR_SAVE_FORMS = ["2024-02-03T00:00:00Z", "2025-06-07T08:09:00Z", "2025-01-02T03:04:05Z"]


def word_results() -> dict[str, str]:
    """Word's field results by token, from ``results.txt``'s ``Dnn_... = `` lines, as docx4j's test reads them."""
    text = (PROBE / "results.txt").read_text(encoding="utf-8")
    return {m.group(1): m.group(2).strip() for m in re.finditer(r"^(D\d\d_[A-Z_]+) = (.*)$", text, re.MULTILINE)}


def results(pkg) -> dict[str, str]:
    """Each token's result in the plain text under ``fields: update``: the line ``D01_... = <result>``."""
    out = {}
    for line in plain_text(extract(pkg, fields="update")).splitlines():
        found = re.match(r"^(D\d\d_[A-Z_]+) = (.*)$", line)
        if found:
            out[found.group(1)] = found.group(2).strip()
    return out


def attribute_values(data: bytes, attribute: str) -> list[str]:
    root = etree.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml"))
    return [value for element in root.iter() if (value := element.get(f"{{{W}}}{attribute}")) is not None]


def last_printed(data: bytes) -> str:
    root = etree.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("docProps/core.xml"))
    return root.find("{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastPrinted").text


def _replay(got: dict[str, str]) -> None:
    """Word's text for CREATEDATE and SAVEDATE; PRINTDATE's cached ``OLD`` (Q6: the text model keeps it)."""
    for token, expected in word_results().items():
        if "PRINTDATE" in token:
            assert got[token] == "OLD", token
        else:
            assert got[token] == expected, token


def test_the_probe_files_are_docx4js():
    sums = dict(reversed(line.split()) for line in (PROBE / "SHA1SUMS").read_text().splitlines())
    assert sorted(sums) == ["README.txt", "date-forms-r1-word365.docx", "date-forms-r1.docx", "results.txt"]
    for name, digest in sums.items():
        assert hashlib.sha1((PROBE / name).read_bytes()).hexdigest() == digest, name


def test_results_txt_holds_the_six_field_cases():
    assert word_results() == {
        "D01_CREATEDATE_DATEONLY": "2024-02-03 11:00:00",
        "D02_SAVEDATE_LOCAL": "2025-06-07 18:09:00",
        "D03_PRINTDATE_UTC": "2025-01-02 14:04:00",
        "D04_CREATEDATE_DEFAULT": "2/3/2024 11:00:00 AM",
        "D05_SAVEDATE_DEFAULT": "6/7/2025 6:09:00 PM",
        "D06_PRINTDATE_DEFAULT": "1/2/2025 2:04:00 PM",
    }


# ---------------------------------------------------------------------------
# Part 1: the fields (item 1)
# ---------------------------------------------------------------------------


def test_the_property_fields_give_words_results_in_the_processs_zone(zone):
    """``thePropertyFieldsGiveWordsResults``, through the text model: D01, D02, D04, D05 Word's; D03, D06 cached."""
    zone(SYDNEY)
    got = results(load(PROBE / "date-forms-r1.docx"))
    _replay(got)
    assert got["D01_CREATEDATE_DATEONLY"] == "2024-02-03 11:00:00", "a date alone is UTC midnight"
    assert got["D05_SAVEDATE_DEFAULT"] == "6/7/2025 6:09:00 PM", "a time with no offset is UTC, to the minute"


@pytest.mark.parametrize("place", ["America/New_York", "Asia/Kolkata", "UTC"])
def test_the_property_fields_give_words_results_in_the_packages_zone(zone, place):
    """docx4j's test sets the package's zone (D12): the process's plays no part."""
    zone(place)
    pkg = load(PROBE / "date-forms-r1.docx")
    pkg.date_time_zone = SYDNEY
    _replay(results(pkg))


def test_a_bare_core_property_value_is_utc(zone):
    """``aBareCorePropertyValueIsUtcAndPrintdateIsLastPrinted``, its D13 half: ``parse_w3cdtf`` and the zone override."""
    zone("UTC")
    assert parse_w3cdtf("2024-02-03") == datetime.datetime(2024, 2, 3, tzinfo=UTC)
    assert parse_w3cdtf("2025-06-07T08:09:10") == datetime.datetime(2025, 6, 7, 8, 9, 10, tzinfo=UTC)
    assert parse_w3cdtf("2025-01-02T03:04:05Z") == datetime.datetime(2025, 1, 2, 3, 4, 5, tzinfo=UTC)
    pkg = load(PROBE / "date-forms-r1.docx")
    pkg.date_time_zone = SYDNEY
    ctx = FieldContext(pkg)
    assert evaluate(' CREATEDATE \\@ "yyyy-MM-dd HH:mm:ss" ', ctx) == "2024-02-03 11:00:00"
    assert evaluate(" SAVEDATE ", ctx) == "6/7/2025 6:09:00 PM"
    tokyo = FieldContext(pkg, zone="Asia/Tokyo")
    assert evaluate(' CREATEDATE \\@ "yyyy-MM-dd HH:mm:ss" ', tokyo) == "2024-02-03 09:00:00"
    assert evaluate(' SAVEDATE \\@ "yyyy-MM-dd HH:mm:ss" ', tokyo) == "2025-06-07 17:09:00"


def test_printdate_here_is_the_registrys_clock_not_last_printed(zone):
    """The other half, docx4j's D14 (``PRINTDATE`` is ``cp:lastPrinted``), is not ported: CR-025 D14 stands (section 13).

    The text model keeps ``PRINTDATE`` cached either way (Q6, D03 and D06
    above); a caller of :func:`evaluate` gets the context's clock, where docx4j
    since ``c58e36e64`` gives ``1/2/2025 2:04:00 PM``.
    """
    zone("UTC")
    pkg = load(PROBE / "date-forms-r1.docx")
    pkg.date_time_zone = SYDNEY
    clock = datetime.datetime(2026, 3, 4, 5, 6, 7, tzinfo=UTC)
    ctx = FieldContext(pkg, clock=lambda: clock)
    assert evaluate(' PRINTDATE \\@ "yyyy-MM-dd HH:mm" ', ctx) == "2026-03-04 16:06"


# ---------------------------------------------------------------------------
# Part 2: the revisions (item 2)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("place", [SYDNEY, "America/New_York"])
def test_the_revisions_are_words_local_clock(zone, place):
    """``theRevisionsAreWordsLocalClock``: R01 local midnight (it was None), R02 as written, R03 the ``Z`` ignored."""
    zone(place)
    pkg = load(PROBE / "date-forms-r1.docx")
    pkg.date_time_zone = SYDNEY
    assert [change.date for change in tracked_changes(pkg)] == WORD_REVISIONS
    assert [change.date for change in pkg.body.get_tracked_changes()] == WORD_REVISIONS


def test_the_revisions_with_the_process_in_sydney_and_no_package_zone(zone):
    zone(SYDNEY)
    assert [change.date for change in tracked_changes(load(PROBE / "date-forms-r1.docx"))] == WORD_REVISIONS


# ---------------------------------------------------------------------------
# Part 3: the date controls
# ---------------------------------------------------------------------------


def test_the_date_controls_are_the_nominal_day():
    """B01 to B03: the picker on the stored day in every form; B01, a date alone, was None before Phase B2."""
    pkg = load(PROBE / "date-forms-r1.docx")
    controls = [c for c in pkg.body.content_controls if c.type == "DatePicker"]
    assert [c.date_picker_content_control.full_date.date() for c in controls] == WORD_PICKER_DAYS
    assert [c.text for c in controls] == ["2024-02-03", "2025-06-07T08:09:10", "2025-01-02T03:04:05Z"]


# ---------------------------------------------------------------------------
# Word's save, and this one's (item 3)
# ---------------------------------------------------------------------------


def test_words_save_writes_the_nominal_clock_to_the_minute():
    """``wordsSaveWritesTheNominalClockToTheMinute``: every ``w:date`` and ``w:fullDate``; ``lastPrinted`` to the minute."""
    data = (PROBE / "date-forms-r1-word365.docx").read_bytes()
    assert attribute_values(data, "date") == WORD_SAVE_FORMS
    assert attribute_values(data, "fullDate") == WORD_SAVE_FORMS
    assert last_printed(data) == "2025-01-02T03:04:00Z"
    assert attribute_values((PROBE / "date-forms-r1.docx").read_bytes(), "date") == [
        "2024-02-03", "2025-06-07T08:09:10", "2025-01-02T03:04:05Z",
    ]


def _saved(**options) -> tuple[object, bytes]:
    pkg = load(PROBE / "date-forms-r1.docx")
    pkg.main_document_part.contents  # noqa: B018 - read, so that the save re-marshals it
    return pkg, pkg.save(**options)


def test_this_save_writes_words_forms_where_the_source_did_not():
    """Item 3: R01, R02, B01 and B02 as Word's save wrote them; R03 and B03 as read, where Word cut ``:05`` to ``:00``."""
    pkg, data = _saved()
    assert attribute_values(data, "date") == OUR_SAVE_FORMS
    assert attribute_values(data, "fullDate") == OUR_SAVE_FORMS
    differs = [(ours, word) for ours, word in zip(OUR_SAVE_FORMS, WORD_SAVE_FORMS, strict=True) if ours != word]
    assert differs == [("2025-01-02T03:04:05Z", "2025-01-02T03:04:00Z")], "the user's decision, and only there"
    assert [(row.path, row.before, row.after) for row in pkg.last_save.normalised] == [
        ("w:document/w:body/w:p[11]/w:ins/@w:date", "2024-02-03", "2024-02-03T00:00:00Z"),
        ("w:document/w:body/w:p[12]/w:ins/@w:date", "2025-06-07T08:09:10", "2025-06-07T08:09:00Z"),
        ("w:document/w:body/w:p[15]/w:sdt/w:sdtPr/w:date/@w:fullDate", "2024-02-03", "2024-02-03T00:00:00Z"),
        ("w:document/w:body/w:p[16]/w:sdt/w:sdtPr/w:date/@w:fullDate", "2025-06-07T08:09:10", "2025-06-07T08:09:00Z"),
    ]
    # the core part is never unmarshalled, so lastPrinted keeps its seconds, where Word's save cut them
    assert last_printed(data) == "2025-01-02T03:04:05Z"


def test_this_save_reads_back_as_words_save_reads(zone):
    """Word's meaning kept: R01 and R02 read the same from this save and Word's; R03 too but for the seconds Word cut."""
    zone(SYDNEY)
    _pkg, data = _saved()
    ours = [change.date for change in tracked_changes(load(io.BytesIO(data)))]
    words = [change.date for change in tracked_changes(load(PROBE / "date-forms-r1-word365.docx"))]
    assert ours[:2] == words[:2] == [WORD_REVISIONS[0], WORD_REVISIONS[1].replace(second=0)]
    assert ours[2] == WORD_REVISIONS[2] and words[2] == WORD_REVISIONS[2].replace(second=0)
    again = load(io.BytesIO(data))
    controls = [c for c in again.body.content_controls if c.type == "DatePicker"]
    assert [c.date_picker_content_control.full_date.date() for c in controls] == WORD_PICKER_DAYS
    assert again.main_document_part.kept == [], "nothing kept as text: every value now binds"


def test_normalise_false_keeps_the_probes_spelling():
    pkg, data = _saved(normalise=False)
    assert pkg.last_save.normalised == []
    assert attribute_values(data, "date") == attribute_values(data, "fullDate") == [
        "2024-02-03", "2025-06-07T08:09:10", "2025-01-02T03:04:05Z",
    ]


def test_canon_counts_the_four_as_word_normal_form_and_nothing_else():
    sys.path.insert(0, str(ROOT / "scripts"))
    import canon

    _pkg, data = _saved()
    source = zipfile.ZipFile(PROBE / "date-forms-r1.docx").read("word/document.xml")
    saved = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml")
    report = canon.compare(source, saved, normalise=True)
    assert report.identical
    assert report.counts == {"word-normal-form": 4}


def test_a_save_that_never_reads_the_document_writes_it_byte_for_byte():
    """The engine's rule: a part never unmarshalled is the source's bytes, its dates included."""
    pkg = load(PROBE / "date-forms-r1.docx")
    data = pkg.save()
    assert zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml") == zipfile.ZipFile(
        PROBE / "date-forms-r1.docx"
    ).read("word/document.xml")
    assert pkg.last_save.normalised == []


def test_the_zone_in_the_probe_is_sydneys():
    """The probe ran at UTC+11 (February, January) and UTC+10 (June): what the expectations above assume."""
    sydney = zoneinfo.ZoneInfo(SYDNEY)
    assert datetime.datetime(2024, 2, 3, tzinfo=sydney).utcoffset() == datetime.timedelta(hours=11)
    assert datetime.datetime(2025, 6, 7, tzinfo=sydney).utcoffset() == datetime.timedelta(hours=10)
    assert datetime.datetime(2025, 1, 2, tzinfo=sydney).utcoffset() == datetime.timedelta(hours=11)
