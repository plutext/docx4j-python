"""A revision's or a comment's stamp as the author's clock. docx4j ``RevisionDates``. CR-027 Phase B.

The port of docx4j's ``org.docx4j.model.revisions.RevisionDates`` at
``docx4j-17.3.5`` (``fc71f12e6``; WD 0.46 REQ-702, the date review's section
3.7 (c)). Word writes two attributes:

* ``w:date``, the author's wall clock, labelled ``Z`` although it is not UTC;
* ``w16du:dateUtc`` (``w16cex:dateUtc`` for a comment), the instant.

Word's Reviewing Pane shows ``w:date`` as it is and ignores ``dateUtc``; a mark
with ``dateUtc`` alone shows no date (measured 2026-10-31, docx4j CR-068 phase
0, ``tests/fixtures/revisions/cr068/manual-redline-dates.docx``). So a stamp
has two readings: the instant
(:attr:`~docx4j_py.model.revisions.change.TrackedChange.date`, CR-025's, read in
the package's zone where there is no ``dateUtc``), and the author's local time,
with the author's UTC offset where both attributes let it be derived. These
are the second. Nothing here converts a time to another zone.

Each function takes the model's ``XmlDateTime``, its lexical form (the exact
parse keeps a date alone as its text), or a ``datetime``
(:func:`~docx4j_py.openpackaging.parts.wml.comments.dates_utc`'s), read by
:func:`docx4j_py.model.values.date_time`; None or anything that is no date is
no date.
"""

from __future__ import annotations

import datetime
from typing import Any

__all__ = ["local_time", "offset", "text"]

#: The offsets clocks keep, in seconds: UTC-12:00 to UTC+14:00.
MIN_OFFSET = -12 * 3600
MAX_OFFSET = 14 * 3600
#: A clock's offset is a whole number of quarter hours.
_QUARTER_HOUR = 15 * 60
_SECOND = datetime.timedelta(seconds=1)


def _clock(value: Any) -> tuple[datetime.datetime, int | None] | None:
    """`value`'s date and time to the second, naive, and its offset in minutes (None for none); None for no date.

    A missing time is 00:00:00 (a date alone is midnight), fractions of a
    second are dropped, and ``24:00:00`` is the next day's midnight, as Java's
    ``XMLGregorianCalendar`` reads it. A year a ``datetime`` cannot hold
    (before 1, after 9999) is no date.
    """
    from docx4j_py.model.values import date_time

    calendar = date_time(value)
    if calendar is None:
        return None
    midnight_after = calendar.hour == 24
    try:
        clock = datetime.datetime(  # noqa: DTZ001 - the author's clock, no zone assumed
            calendar.year,
            calendar.month,
            calendar.day,
            0 if midnight_after else calendar.hour,
            calendar.minute,
            calendar.second,
        )
        if midnight_after:
            clock += datetime.timedelta(days=1)
    except (ValueError, OverflowError):
        return None
    return clock, calendar.offset


def local_time(date: Any) -> datetime.datetime | None:
    """The author's local time: ``w:date``'s date and time as stored, naive. docx4j ``RevisionDates.localTime``.

    Any designator (a ``Z``, an offset) is not read, a missing time is
    00:00:00, and fractions of a second are dropped; None for no date.
    """
    found = _clock(date)
    return None if found is None else found[0]


def _utc(date_utc: Any) -> datetime.datetime | None:
    """``dateUtc`` as a UTC clock, naive: one with no designator is UTC, an explicit offset is applied; None for none."""
    found = _clock(date_utc)
    if found is None:
        return None
    clock, minutes = found
    return clock if minutes is None else clock - datetime.timedelta(minutes=minutes)


def offset(date: Any, date_utc: Any) -> datetime.timezone | None:
    """The author's UTC offset: ``w:date`` less ``dateUtc``. docx4j ``RevisionDates.offset``.

    It counts only where both are present and differ by a whole number of
    quarter hours from -12:00 to +14:00, as a clock's offset does; else None.
    A pair written inconsistently gives no offset rather than a wrong one.
    UTC is :data:`datetime.UTC`.
    """
    local, utc = local_time(date), _utc(date_utc)
    if local is None or utc is None:
        return None
    seconds = (local - utc) // _SECOND
    if seconds % _QUARTER_HOUR != 0 or seconds < MIN_OFFSET or seconds > MAX_OFFSET:
        return None
    return datetime.UTC if seconds == 0 else datetime.timezone(datetime.timedelta(seconds=seconds))


def _local_text(clock: datetime.datetime) -> str:
    """ISO 8601's extended form with no designator: ``2026-10-31T11:00:00`` (Java's ``uuuu-MM-dd'T'HH:mm:ss``)."""
    return clock.isoformat(timespec="seconds")


def _offset_text(zone: datetime.timezone) -> str:
    """``+11:00``, ``-05:00``, ``+00:00`` (Java's ``xxx``), never ``Z``."""
    seconds = zone.utcoffset(None) // _SECOND
    sign = "-" if seconds < 0 else "+"
    hours, minutes = divmod(abs(seconds) // 60, 60)
    return f"{sign}{hours:02d}:{minutes:02d}"


def text(date: Any, date_utc: Any) -> str | None:
    """The stamp as REQ-702 writes it. docx4j ``RevisionDates.text``.

    * ``w:date`` with an offset known: the local time with the offset after it
      (``2026-10-31T11:00:00+11:00``; ``+00:00`` for UTC, never ``Z``);
    * ``w:date`` alone: the local time with no designator (``2026-10-31T11:00:00``);
    * ``dateUtc`` alone: written with ``Z`` (``2026-10-31T00:00:00Z``);
    * neither: None.

    The ``Z`` Word stores on ``w:date`` is never written, and no time is
    converted to another zone.
    """
    local = local_time(date)
    if local is None:
        utc = _utc(date_utc)
        return None if utc is None else _local_text(utc) + "Z"
    zone = offset(date, date_utc)
    return _local_text(local) if zone is None else _local_text(local) + _offset_text(zone)
