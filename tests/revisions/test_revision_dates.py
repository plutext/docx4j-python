"""A stamp as the author's clock: docx4j ``RevisionDatesTest``, ported whole. CR-027 Phase B (sections 2.2, 4.2, 5).

docx4j's ``RevisionDatesTest`` at ``docx4j-17.3.5`` (``fc71f12e6``; WD 0.46
REQ-702), over :mod:`docx4j_py.model.revisions.dates`:

- ``measuredCases``: REQ-702's form for Word's four measured cases, and for no date;
- ``offsets``: any offset a clock keeps, west of UTC, half and three-quarter
  hours, and none for an inconsistent pair;
- ``otherForms``: ``w:date``'s own designator never read, ``dateUtc``'s offset
  applied, a date alone midnight;
- ``trackedChanges``: :attr:`TrackedChange.author_local_time` and
  :attr:`~TrackedChange.author_offset` beside :attr:`~TrackedChange.date`, on
  ``manual-redline-dates.docx`` (``tests/fixtures/revisions/cr068/``, docx4j's
  ``revisions/cr068/`` copied at the tag, CR-028 Q11).

docx4j's ``x(s)`` is an ``XMLGregorianCalendar``; here each case runs over the
model's form (an ``XmlDateTime``, or the text the exact parse keeps for a date
alone) and over the lexical form. Then this port's own: the copy held to
``SHA1SUMS`` and to docx4j's blobs, what is no date, ``24:00:00``, the
``datetime`` that ``dates_utc`` gives with its microseconds, a group's stamp,
and the instant (CR-025's) unchanged.
"""

from __future__ import annotations

import datetime
import hashlib
import subprocess
from pathlib import Path

import pytest
from revision_helpers import run, with_body

from docx4j_py import load
from docx4j_py.model import revisions
from docx4j_py.model.revisions import dates
from docx4j_py.model.revisions.dates import local_time, offset

ROOT = Path(__file__).resolve().parent.parent.parent
CR068 = ROOT / "tests" / "fixtures" / "revisions" / "cr068"
DOCX4J = ROOT.parent / "docx4j"
TAG = "docx4j-17.3.5"
SOURCE = "docx4j-core-tests/src/test/resources/revisions/cr068"
UTC = datetime.UTC


def _model(s: str):
    """``s`` as the model holds it: an ``XmlDateTime``, or the text the exact parse keeps (a date alone)."""
    from docx4j_xsdata.models.datatype import XmlDateTime

    try:
        return XmlDateTime.from_string(s)
    except ValueError:
        return s


@pytest.fixture(params=["model", "lexical"])
def x(request):
    """docx4j's ``x(s)``: a stamp in the model's form, or in its lexical form."""
    return _model if request.param == "model" else str


@pytest.fixture
def text(x):
    """docx4j's ``text(date, dateUtc)``, None for either left out."""

    def text(date: str | None, date_utc: str | None) -> str | None:
        return dates.text(None if date is None else x(date), None if date_utc is None else x(date_utc))

    return text


def zone(hours: int, minutes: int = 0) -> datetime.timezone:
    """docx4j's ``ZoneOffset.ofHoursMinutes``."""
    return datetime.timezone(datetime.timedelta(hours=hours, minutes=minutes))


# ---------------------------------------------------------------------------
# RevisionDatesTest
# ---------------------------------------------------------------------------


def test_measured_cases(text):
    """``measuredCases``: REQ-702's form for Word's four measured cases, and for no date."""
    assert text("2026-10-31T00:00:00Z", None) == "2026-10-31T00:00:00"  # D1: the Z not written
    assert text(None, "2026-10-31T00:00:00Z") == "2026-10-31T00:00:00Z"  # D2: dateUtc alone
    assert text("2026-10-31T00:00:00Z", "2026-10-31T00:00:00Z") == "2026-10-31T00:00:00+00:00"  # D3: an author in UTC
    assert (
        text("2026-10-31T11:00:00Z", "2026-10-31T00:00:00Z") == "2026-10-31T11:00:00+11:00"
    )  # D4: Sydney, as Word writes
    assert text(None, None) is None


def test_offsets(text, x):
    """``offsets``: any offset a clock keeps; an inconsistent pair gives none."""
    assert text("2026-10-30T19:00:00Z", "2026-10-31T00:00:00Z") == "2026-10-30T19:00:00-05:00"
    assert text("2026-10-31T05:30:00Z", "2026-10-31T00:00:00Z") == "2026-10-31T05:30:00+05:30"
    assert text("2026-10-31T05:45:00Z", "2026-10-31T00:00:00Z") == "2026-10-31T05:45:00+05:45"
    assert text("2026-10-31T14:00:00Z", "2026-10-31T00:00:00Z") == "2026-10-31T14:00:00+14:00"
    assert text("2026-10-30T12:00:00Z", "2026-10-31T00:00:00Z") == "2026-10-30T12:00:00-12:00"
    # not a quarter hour, or beyond any clock: no offset rather than a wrong one
    assert text("2026-10-31T00:07:00Z", "2026-10-31T00:00:00Z") == "2026-10-31T00:07:00"
    assert text("2026-11-01T03:00:00Z", "2026-10-31T00:00:00Z") == "2026-11-01T03:00:00"
    assert text("2026-10-30T11:00:00Z", "2026-10-31T00:00:00Z") == "2026-10-30T11:00:00"
    assert offset(x("2026-10-31T00:07:00Z"), x("2026-10-31T00:00:00Z")) is None
    assert offset(x("2026-10-31T05:45:00Z"), x("2026-10-31T00:00:00Z")) == zone(5, 45)


def test_other_forms(text, x):
    """``otherForms``: ``w:date``'s own designator is never read; ``dateUtc``'s offset is applied; a date alone is midnight."""
    assert text("2026-10-31T11:00:00", None) == "2026-10-31T11:00:00"
    assert text("2026-10-31T11:00:00+05:00", None) == "2026-10-31T11:00:00"
    assert text("2026-10-31", None) == "2026-10-31T00:00:00"
    assert text("2026-10-31T11:00:00.5Z", None) == "2026-10-31T11:00:00"
    assert text("2026-10-31T11:00:00Z", "2026-10-31T05:00:00+05:00") == "2026-10-31T11:00:00+11:00"
    assert text(None, "2026-10-31T05:00:00+05:00") == "2026-10-31T00:00:00Z"
    assert local_time(x("2026-10-31T11:00:00Z")) == datetime.datetime(2026, 10, 31, 11, 0)  # noqa: DTZ001 - the author's clock is naive
    assert local_time(None) is None


def test_tracked_changes():
    """``trackedChanges``: the author's local time and offset beside the instant, on Word's four cases."""
    pkg = load(CR068 / "manual-redline-dates.docx")
    changes = revisions.tracked_changes(pkg)
    assert len(changes) == 4
    d1, d2, d3, d4 = changes
    assert d1.author_local_time == datetime.datetime(2026, 10, 31, 0, 0)  # noqa: DTZ001
    assert d1.author_offset is None
    assert d2.author_local_time is None  # Word shows no date
    assert d2.author_offset is None
    assert d2.date == datetime.datetime(2026, 10, 31, tzinfo=UTC)
    assert d3.author_local_time == datetime.datetime(2026, 10, 31, 0, 0)  # noqa: DTZ001
    assert d3.author_offset == UTC
    assert d4.author_local_time == datetime.datetime(2026, 10, 31, 11, 0)  # noqa: DTZ001
    assert d4.author_offset == zone(11)
    # where the offset is known, the author's clock gives the instant, with no zone assumed
    assert d4.date == d4.author_local_time.replace(tzinfo=d4.author_offset)
    assert d3.date == d3.author_local_time.replace(tzinfo=d3.author_offset)


# ---------------------------------------------------------------------------
# the copy (CR-028 Q11)
# ---------------------------------------------------------------------------


def _sums() -> dict[str, str]:
    return dict(reversed(line.split("  ", 1)) for line in (CR068 / "SHA1SUMS").read_text(encoding="utf-8").splitlines())


def test_the_copy_is_every_file_sha1sums_lists():
    """Byte for byte, 54 files (docx4j's 58 less the four screenshots), and nothing else beside them but the README."""
    listed = _sums()
    for name, digest in listed.items():
        assert hashlib.sha1((CR068 / name).read_bytes()).hexdigest() == digest, name
    assert sorted(p.name for p in CR068.iterdir()) == sorted([*listed, "SHA1SUMS", "README.md"])
    assert len(listed) == 54
    assert not any(name.lower().endswith(".png") for name in listed)


@pytest.mark.skipif(not (DOCX4J / ".git").exists(), reason="needs the sibling checkout ../docx4j")
def test_each_file_is_docx4js_at_the_tag():
    """Each file's git blob id is docx4j's at the tag; the tag's other files are the four screenshots."""
    listing = subprocess.run(
        ["git", "-C", str(DOCX4J), "ls-tree", f"{TAG}:{SOURCE}"], capture_output=True, check=True, text=True
    ).stdout
    blobs = {name: line.split()[2] for line in listing.splitlines() for name in [line.split("\t", 1)[1]]}
    for name in _sums():
        data = (CR068 / name).read_bytes()
        assert hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest() == blobs.pop(name), name
    assert sorted(blobs) == [
        "cr068-comment-cards-pane.PNG",
        "cr068-comment-cards.PNG",
        "cr068-default.PNG",
        "cr068-optins.png",
    ]


# ---------------------------------------------------------------------------
# this port's own
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("value", ["", "  ", "abc", "2025-06-07T08:09", "11:00:00", 42])
def test_what_is_no_date(value):
    """A value that is no date is none, as JAXB gives docx4j null; and none alone is no stamp."""
    assert local_time(value) is None
    assert dates.text(value, None) is None
    assert dates.text(value, value) is None
    assert offset("2026-10-31T11:00:00Z", value) is None


def test_twenty_four_hundred_is_the_next_days_midnight():
    """``24:00:00``, which ``xsd:dateTime`` allows, is the next day's midnight, as Java's ``XMLGregorianCalendar`` reads it."""
    assert local_time("2026-10-31T24:00:00Z") == datetime.datetime(2026, 11, 1)  # noqa: DTZ001
    assert dates.text("2026-11-01T11:00:00Z", "2026-10-31T24:00:00Z") == "2026-11-01T11:00:00+11:00"


def test_a_datetime_is_read_to_the_second():
    """``dates_utc``'s ``datetime`` keeps microseconds; ``RevisionDates`` reads ``dateUtc`` to the second (section 4.2)."""
    half = datetime.datetime(2026, 10, 31, 0, 0, 0, 500000, tzinfo=UTC)
    assert dates.text("2026-10-31T11:00:00Z", half) == "2026-10-31T11:00:00+11:00"
    assert dates.text(None, half) == "2026-10-31T00:00:00Z"
    assert dates.text("2026-10-31T11:00:00Z", "2026-10-31T00:00:00.5Z") == "2026-10-31T11:00:00+11:00"
    # an aware datetime in another zone is applied, as an explicit offset is
    assert dates.text("2026-10-31T11:00:00Z", half.astimezone(zone(5, 30))) == "2026-10-31T11:00:00+11:00"


def test_a_groups_stamp_is_its_first_pieces():
    """docx4j: "A group's is its first piece's": touching deletions by one author are one change."""
    package = with_body(
        '<w:p><w:del w:id="1" w:author="A" w:date="2026-10-31T11:00:00Z" w16du:dateUtc="2026-10-31T00:00:00Z">'
        "<w:r><w:delText>one </w:delText></w:r></w:del>"
        '<w:del w:id="2" w:author="A" w:date="2026-10-31T09:30:00Z">'
        "<w:r><w:delText>two</w:delText></w:r></w:del>" + run(" kept") + "</w:p>"
    )
    (change,) = revisions.tracked_changes(package)
    assert len(change.pieces) == 2
    assert change.author_local_time == datetime.datetime(2026, 10, 31, 11, 0)  # noqa: DTZ001
    assert change.author_offset == zone(11)
    assert change.pieces[1].author_local_time == datetime.datetime(2026, 10, 31, 9, 30)  # noqa: DTZ001
    assert change.pieces[1].author_offset is None


def test_the_instant_is_cr025s_reading_unchanged():
    """``date`` is the instant as it was (Q1): ``dateUtc``, else ``w:date`` in the package's zone; the content API's too."""
    pkg = load(CR068 / "manual-redline-dates.docx")
    pkg.date_time_zone = "Australia/Sydney"
    d1, d2, d3, d4 = revisions.tracked_changes(pkg)
    assert d1.date == datetime.datetime(2026, 10, 30, 13, 0, tzinfo=UTC)  # w:date alone, read at UTC+11
    assert d1.author_local_time == datetime.datetime(2026, 10, 31, 0, 0)  # noqa: DTZ001 - the author's clock, no zone assumed
    assert [c.date for c in (d2, d3, d4)] == [datetime.datetime(2026, 10, 31, tzinfo=UTC)] * 3
    assert [c.date for c in pkg.get_tracked_changes()] == [c.date for c in (d1, d2, d3, d4)]
