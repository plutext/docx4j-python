"""The package's time zone: where an instant meets a clock. CR-025 Phase A.

The portfolio's date rule (``docx4j-portfolio/docs/date_handling_across_ports.md``
section 3.6, the user's ruling of 2026-10-28), after docx4j CR-067 D12
(``OpcPackage.setDateTimeZone``, ``effectiveDateTimeZone()``):

* a revision's and a comment's stamp, and a date field, hold an **instant**; a
  bound date holds a nominal day; an extractor gives every stored date as
  written and converts nothing;
* an instant is shown, and a ``w:date`` without ``dateUtc`` read, in the
  **package's zone** (``pkg.date_time_zone``), the client's for this package;
  unset, in the process's (``TZ``), and nothing between (CR-025 D2);
* a new revision or comment is stamped to the minute, as Word writes it.

Word writes a ``w:date`` as the author's local wall-clock time labelled ``Z``
(CR-010 section 12.9), so the zone is what turns it into an instant and back.
Every function here takes ``zone=None`` to mean the process's zone, and then
goes the way the code went before the package had a zone (``astimezone()``
with no argument), so that a package with none set writes and reads exactly
what it did. :data:`PROCESS_ZONE` is that zone as an object, for
``pkg.effective_date_time_zone``.
"""

from __future__ import annotations

import datetime
import time
import zoneinfo
from typing import Any

__all__ = [
    "PROCESS_ZONE",
    "ProcessZone",
    "effective_zone",
    "instant_of",
    "package_zone",
    "to_the_minute",
    "wall_clock",
    "zone_of",
]

_UTC = datetime.UTC


class ProcessZone(datetime.tzinfo):
    """The process's zone (``TZ``), read at each instant exactly as ``datetime.astimezone()`` reads it.

    The standard library has no object for the local zone: ``astimezone()``
    with no argument gives a fixed offset for the one instant. This one asks
    again at every instant, so it holds across a daylight-saving change, and
    a wall-clock time in a fold is told apart by its ``fold``, as PEP 495
    says. Equal only to itself; :data:`PROCESS_ZONE` is the one to use.
    """

    def utcoffset(self, dt: datetime.datetime | None) -> datetime.timedelta:
        """The offset of the wall-clock time `dt` (its ``fold`` read), or of now."""
        local = _local(dt)
        offset = local.utcoffset()
        assert offset is not None
        return offset

    def dst(self, dt: datetime.datetime | None) -> datetime.timedelta:
        """The daylight-saving part of :meth:`utcoffset`, against the zone's standard offset."""
        local = _local(dt)
        if time.localtime(local.timestamp()).tm_isdst <= 0:
            return datetime.timedelta(0)
        offset = local.utcoffset()
        assert offset is not None
        return offset - datetime.timedelta(seconds=-time.timezone)

    def tzname(self, dt: datetime.datetime | None) -> str | None:
        """The zone's abbreviation at `dt` (``AEDT``), as the platform gives it."""
        return _local(dt).tzname()

    def fromutc(self, dt: datetime.datetime) -> datetime.datetime:
        """The wall-clock time the UTC time `dt` shows here, its ``fold`` set where the time comes twice."""
        if not isinstance(dt, datetime.datetime) or dt.tzinfo is not self:
            raise ValueError("fromutc: dt.tzinfo is not self")
        local = dt.replace(tzinfo=_UTC).astimezone()
        naive = local.replace(tzinfo=None)
        fold = 0 if naive.replace(fold=0).astimezone().utcoffset() == local.utcoffset() else 1
        return naive.replace(tzinfo=self, fold=fold)

    def __repr__(self) -> str:
        """``ProcessZone()``."""
        return "ProcessZone()"

    def __reduce__(self) -> str:
        """Pickled as the one instance."""
        return "PROCESS_ZONE"


def _local(dt: datetime.datetime | None) -> datetime.datetime:
    """`dt` read as a wall-clock time in the process's zone (now for None), aware with that instant's offset."""
    if dt is None:
        return datetime.datetime.now().astimezone()
    return dt.replace(tzinfo=None).astimezone()


#: The process's zone, the one :class:`ProcessZone`.
PROCESS_ZONE = ProcessZone()


def _is_process(zone: datetime.tzinfo | None) -> bool:
    return zone is None or isinstance(zone, ProcessZone)


def zone_of(value: Any) -> datetime.tzinfo | None:
    """A zone from what a caller gives: None, a ``datetime.tzinfo``, or an IANA name (``"Asia/Tokyo"``).

    Raises:
        ValueError: a name the platform's zone database does not know (``ZoneInfo`` raises).
        TypeError: anything else.
    """
    if value is None or isinstance(value, datetime.tzinfo):
        return value
    if isinstance(value, str):
        name = value.strip()
        if name:
            try:
                return zoneinfo.ZoneInfo(name)
            except (zoneinfo.ZoneInfoNotFoundError, ValueError, OSError):
                pass
        raise ValueError(f"not a time zone: {value!r} (an IANA name, such as 'Australia/Sydney')")
    raise TypeError(f"a time zone is a datetime.tzinfo or an IANA name, not {type(value).__name__}")


def package_zone(package: Any) -> datetime.tzinfo | None:
    """The zone `package` names (``pkg.date_time_zone``), or None for the process's; None for no package."""
    if package is None:
        return None
    try:
        zone = getattr(package, "date_time_zone", None)
    except Exception:  # noqa: BLE001 - a package-like object with no zone shows the process's
        return None
    return zone if isinstance(zone, datetime.tzinfo) else None


def effective_zone(package: Any) -> datetime.tzinfo:
    """The zone an instant in `package` is shown in: its own, else :data:`PROCESS_ZONE`. docx4j ``effectiveDateTimeZone()``."""
    zone = package_zone(package)
    return zone if zone is not None else PROCESS_ZONE


def wall_clock(instant: datetime.datetime, zone: datetime.tzinfo | None = None) -> datetime.datetime:
    """The wall-clock time, naive, that `instant` shows in `zone` (None: the process's). A naive `instant` is UTC."""
    aware = instant if instant.tzinfo is not None else instant.replace(tzinfo=_UTC)
    local = aware.astimezone() if _is_process(zone) else aware.astimezone(zone)
    return local.replace(tzinfo=None)


def instant_of(clock: datetime.datetime, zone: datetime.tzinfo | None = None) -> datetime.datetime:
    """The instant, in UTC, at which a wall clock in `zone` (None: the process's) shows `clock`; its own tzinfo ignored."""
    naive = clock.replace(tzinfo=None)
    if _is_process(zone):
        # a naive datetime is the process's local time to astimezone()
        return naive.astimezone(_UTC)
    return naive.replace(tzinfo=zone).astimezone(_UTC)


def to_the_minute(instant: datetime.datetime) -> datetime.datetime:
    """`instant` in UTC with its seconds and microseconds dropped: Word's stamp. A naive `instant` is UTC."""
    aware = instant if instant.tzinfo is not None else instant.replace(tzinfo=_UTC)
    return aware.astimezone(_UTC).replace(second=0, microsecond=0)
