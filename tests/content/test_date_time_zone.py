"""The package's time zone: ``pkg.date_time_zone``, and every place an instant meets a clock. CR-025 Phase A.

The portfolio's date rule (``docx4j-portfolio/docs/date_handling_across_ports.md``
section 3.6, 2026-10-28) after docx4j CR-067 D12 (``c3e9920e5``). docx4j's
``PackageDateTimeZoneTest``, the cases that apply: the fallback chain (the
package's zone, else the process's; docx4j's JVM-wide property between the two
has no counterpart, D2), and a revision's ``w:date`` read in the package's zone
with its ``dateUtc`` unchanged. Its date-field cases are Phase B's
(``tests/text/test_field_probe.py``); its bound-date case waits for CR-014.
Then the writers: a revision's and a comment's ``w:date`` on the package's
wall clock, and with no zone set exactly the bytes the process's zone gave
before (CR-010 section 12.9).
"""

from __future__ import annotations

import datetime
import re
import zoneinfo

import pytest
from conftest import fixture

from docx4j_py import create_package
from docx4j_py.fragments import wml
from docx4j_py.model.content import Author
from docx4j_py.model.content.errors import ContentError
from docx4j_py.model.dates import (
    PROCESS_ZONE,
    ProcessZone,
    effective_zone,
    instant_of,
    package_zone,
    wall_clock,
)

UTC = datetime.UTC
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
W16DU = 'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"'


def _with_paragraph(xml: str):
    package = create_package()
    package.main_document_part.contents.body.content.append(wml(xml))
    return package


# ---------------------------------------------------------------------------
# PackageDateTimeZoneTest.theFallbackChain
# ---------------------------------------------------------------------------


def test_the_fallback_chain(zone):
    """None by default, the process's zone in force; the package's own wins; None returns to the process's."""
    zone("Australia/Sydney")
    pkg = create_package()
    assert pkg.date_time_zone is None
    assert pkg.effective_date_time_zone is PROCESS_ZONE
    pkg.date_time_zone = "America/New_York"
    assert pkg.date_time_zone == zoneinfo.ZoneInfo("America/New_York")
    assert pkg.effective_date_time_zone == zoneinfo.ZoneInfo("America/New_York"), "the package's own wins"
    tokyo = datetime.timezone(datetime.timedelta(hours=9), "JST")
    pkg.date_time_zone = tokyo
    assert pkg.date_time_zone is tokyo, "any tzinfo, as given"
    pkg.date_time_zone = None
    assert pkg.effective_date_time_zone is PROCESS_ZONE
    assert effective_zone(None) is PROCESS_ZONE and package_zone(None) is None


@pytest.mark.parametrize("bad", ["not a zone", "", "  ", "../etc/passwd", 9, datetime.timedelta(hours=9)])
def test_a_zone_the_platform_does_not_know_is_refused_at_the_set(bad):
    """A wrong zone shifts every stamp, so it is refused rather than ignored (docx4j-core-ts's ``RangeError``)."""
    pkg = create_package()
    pkg.date_time_zone = "Asia/Tokyo"
    with pytest.raises(ContentError) as raised:
        pkg.date_time_zone = bad
    assert raised.value.code == "date_time_zone.invalid"
    assert pkg.date_time_zone == zoneinfo.ZoneInfo("Asia/Tokyo"), "the zone set before stays"


def test_the_zone_is_never_saved_and_a_trial_reads_its_packages():
    """A request attribute, not a document one (docx4j D12): nothing of it reaches the file."""
    pkg = create_package()
    pkg.body.insert_paragraph("Text.")
    before = pkg.save()
    pkg.date_time_zone = "Asia/Tokyo"
    after = pkg.save()
    import io
    import zipfile

    def parts(data: bytes) -> dict[str, bytes]:
        archive = zipfile.ZipFile(io.BytesIO(data))
        return {name: archive.read(name) for name in archive.namelist()}

    assert parts(before) == parts(after)
    with pkg.dry_run() as trial:
        assert trial.date_time_zone == zoneinfo.ZoneInfo("Asia/Tokyo")
        assert trial.effective_date_time_zone == zoneinfo.ZoneInfo("Asia/Tokyo")


# ---------------------------------------------------------------------------
# PackageDateTimeZoneTest.aRevisionsLocalClockIsReadInThePackagesZone
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("place", ["Australia/Sydney", "America/New_York"])
def test_a_revisions_local_clock_is_read_in_the_packages_zone(zone, place):
    """``w:date`` is Word's local clock, here Tokyo's: 04:05 there is 19:05Z the day before; ``dateUtc`` is UTC."""
    from docx4j_py.model.revisions import tracked_changes

    zone(place)  # the process's zone plays no part once the package names one
    pkg = _with_paragraph(
        f'<w:p {NS}><w:ins w:id="1" w:author="A" w:date="2024-02-03T04:05:06Z"><w:r><w:t>x</w:t></w:r></w:ins></w:p>'
    )
    pkg.date_time_zone = "Asia/Tokyo"
    assert [c.date for c in tracked_changes(pkg)] == [datetime.datetime(2024, 2, 2, 19, 5, 6, tzinfo=UTC)]
    assert [c.date for c in pkg.body.get_tracked_changes()] == [datetime.datetime(2024, 2, 2, 19, 5, 6, tzinfo=UTC)]
    pkg.date_time_zone = "UTC"
    assert [c.date for c in tracked_changes(pkg)] == [datetime.datetime(2024, 2, 3, 4, 5, 6, tzinfo=UTC)]

    pkg2 = _with_paragraph(
        f'<w:p {NS} {W16DU}><w:ins w:id="1" w:author="A" w:date="2024-02-03T04:05:06Z" '
        'w16du:dateUtc="2024-02-02T19:05:06Z"><w:r><w:t>x</w:t></w:r></w:ins></w:p>'
    )
    pkg2.date_time_zone = "Pacific/Honolulu"
    assert [c.date for c in pkg2.body.get_tracked_changes()] == [datetime.datetime(2024, 2, 2, 19, 5, 6, tzinfo=UTC)]


def test_a_revisions_local_clock_with_no_zone_set_is_the_processs(zone):
    """Unset, the reading is CR-010 section 12.9's: the process's wall clock."""
    zone("Australia/Sydney")  # AEDT, UTC+11, in February
    pkg = _with_paragraph(
        f'<w:p {NS}><w:ins w:id="1" w:author="A" w:date="2024-02-03T04:05:06Z"><w:r><w:t>x</w:t></w:r></w:ins></w:p>'
    )
    assert pkg.body.get_tracked_changes()[0].date == datetime.datetime(2024, 2, 2, 17, 5, 6, tzinfo=UTC)
    zone("America/New_York")  # EST, UTC-5
    assert pkg.body.get_tracked_changes()[0].date == datetime.datetime(2024, 2, 3, 9, 5, 6, tzinfo=UTC)


def test_an_explicit_offset_is_read_as_it_says_whatever_the_zone():
    """Word never writes one; one that is not zero is an instant already (CR-010 section 12.9)."""
    pkg = _with_paragraph(
        f'<w:p {NS}><w:ins w:id="1" w:author="A" w:date="2024-02-03T04:05:06+05:00"><w:r><w:t>x</w:t></w:r></w:ins></w:p>'
    )
    pkg.date_time_zone = "Asia/Tokyo"
    assert pkg.body.get_tracked_changes()[0].date == datetime.datetime(2024, 2, 2, 23, 5, 6, tzinfo=UTC)


def test_a_word_2010_comments_date_is_read_in_the_packages_zone(zone):
    """``comments.docx``'s ``w:date="2011-02-01T11:52:00Z"`` (no ``w16cex``): Tokyo's 11:52 is 02:52Z."""
    zone("America/New_York")
    package = fixture("comments.docx")
    package.date_time_zone = "Asia/Tokyo"
    assert package.body.get_comments()[0].creation_date == datetime.datetime(2011, 2, 1, 2, 52, tzinfo=UTC)
    package.date_time_zone = None
    assert package.body.get_comments()[0].creation_date == datetime.datetime(2011, 2, 1, 16, 52, tzinfo=UTC)


@pytest.mark.parametrize("place", ["Australia/Sydney", "America/New_York"])
def test_a_word_365_comments_w16cex_date_utc_is_utc_whatever_the_packages_zone(zone, place):
    zone(place)
    package = fixture("comments-modern.docx")
    package.date_time_zone = "Asia/Tokyo"
    comment = next(c for c in package.body.get_comments() if c.author_name == "Buck Cronk")
    assert comment.creation_date == datetime.datetime(2026, 5, 18, 23, 21, tzinfo=UTC)


# ---------------------------------------------------------------------------
# the writers
# ---------------------------------------------------------------------------

WHEN = datetime.datetime(2026, 9, 17, 12, 0, tzinfo=UTC)


def _dates(xml: str) -> list[tuple[str, str | None]]:
    return [
        (m.group(1), m.group(2))
        for m in re.finditer(r'<w:(?:ins|del) [^>]*?w:date="([^"]*)"(?: w16du:dateUtc="([^"]*)")?', xml)
    ]


def _tracked():
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude", initials="C")
    package.tracked_change_date = WHEN
    paragraph = package.body.insert_paragraph("The quick brown fox.")
    package.change_tracking_mode = "TrackAll"
    return package, paragraph


@pytest.mark.parametrize("place", ["Australia/Sydney", "America/New_York"])
def test_a_revision_is_stamped_on_the_packages_wall_clock(zone, place):
    """``w:date`` on Tokyo's clock (UTC+9), ``w16du:dateUtc`` the true UTC, whatever the process's zone."""
    zone(place)
    package, paragraph = _tracked()
    package.date_time_zone = "Asia/Tokyo"
    paragraph.insert_text("New. ", location="Start")
    assert _dates(paragraph.get_xml()) == [("2026-09-17T21:00:00Z", "2026-09-17T12:00:00Z")]
    assert package.body.get_tracked_changes()[0].date == WHEN, "it reads back as the instant written"


@pytest.mark.parametrize(("place", "local"), [("Australia/Sydney", "2026-09-17T22:00:00Z"), ("America/New_York", "2026-09-17T08:00:00Z")])
def test_unset_a_revision_is_stamped_on_the_processs_wall_clock_as_before(zone, place, local):
    """CR-010 section 12.9's bytes, unchanged."""
    zone(place)
    _, paragraph = _tracked()
    paragraph.insert_text("New. ", location="Start")
    assert _dates(paragraph.get_xml()) == [(local, "2026-09-17T12:00:00Z")]


def test_a_new_comment_is_stamped_on_the_packages_wall_clock(zone, monkeypatch):
    """The comment writer's "now", to the minute, on Tokyo's clock in ``w:date``, whatever the process's zone."""
    from docx4j_py.model.content import tracking

    zone("America/New_York")
    monkeypatch.setattr(tracking, "now_to_the_minute", lambda: WHEN)
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude", initials="C")
    package.date_time_zone = "Asia/Tokyo"
    package.body.insert_paragraph("The quick brown fox.")
    comment = package.body.paragraphs[0].insert_comment("A note.")
    written = re.search(r"<w:comment [^>]*>", package.main_document_part.comments_part.get_xml()).group(0)
    assert 'w:date="2026-09-17T21:00:00Z"' in written
    assert package.main_document_part.comments_extensible_part is None, "a created package has no w16cex part"
    assert comment.creation_date == WHEN, "w:date alone, read back on the package's clock"
    package.date_time_zone = None
    # w:date alone again, now on New York's clock: Tokyo's 21:00 read there is 01:00Z the next day
    assert package.body.get_comments()[0].creation_date == datetime.datetime(2026, 9, 18, 1, 0, tzinfo=UTC)


# ---------------------------------------------------------------------------
# the process's zone as an object, held to astimezone()
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("place", ["Australia/Sydney", "America/New_York", "Australia/Lord_Howe"])
def test_the_process_zone_is_astimezone_at_every_instant(zone, place):
    """Every hour and half hour of 2026, both ways, gaps and folds included (CR-025 risk 2)."""
    zone(place)
    start = datetime.datetime(2026, 1, 1, tzinfo=UTC)
    for k in range(365 * 48):
        instant = start + datetime.timedelta(minutes=30 * k)
        mine = instant.astimezone(PROCESS_ZONE)
        theirs = instant.astimezone()
        assert mine.replace(tzinfo=None) == theirs.replace(tzinfo=None), instant
        assert mine.utcoffset() == theirs.utcoffset(), instant
        assert mine.astimezone(UTC) == instant, instant
        assert wall_clock(instant, PROCESS_ZONE) == wall_clock(instant) == theirs.replace(tzinfo=None)
        naive = theirs.replace(tzinfo=None)
        for fold in (0, 1):
            clock = naive.replace(fold=fold)
            assert clock.replace(tzinfo=PROCESS_ZONE).astimezone(UTC) == clock.astimezone(UTC), (clock, fold)
            assert instant_of(clock, PROCESS_ZONE) == instant_of(clock) == clock.astimezone(UTC)
    assert isinstance(PROCESS_ZONE, ProcessZone) and repr(PROCESS_ZONE) == "ProcessZone()"
    import pickle

    assert pickle.loads(pickle.dumps(PROCESS_ZONE)) is PROCESS_ZONE


def test_xml_date_unset_is_the_old_formula(zone):
    """``tracking.xml_date`` with no zone: ``instant.astimezone()`` on the process's clock, as it was."""
    from docx4j_xsdata.models.datatype import XmlDateTime

    from docx4j_py.model.content.tracking import xml_date

    for place in ("Australia/Sydney", "America/New_York", "UTC"):
        zone(place)
        for month in range(1, 13):
            instant = datetime.datetime(2026, month, 15, 3, 4, 5, 678, tzinfo=UTC)
            old = XmlDateTime.from_datetime(instant.astimezone().replace(tzinfo=UTC, microsecond=0))
            assert xml_date(instant) == old
            assert xml_date(instant, zoneinfo.ZoneInfo("Asia/Tokyo")) == XmlDateTime.from_datetime(
                instant.astimezone(zoneinfo.ZoneInfo("Asia/Tokyo")).replace(tzinfo=UTC, microsecond=0)
            )
