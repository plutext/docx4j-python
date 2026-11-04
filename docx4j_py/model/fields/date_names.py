"""The month, weekday and AM/PM names a date is written in. CR-014 phase 1b.

docx4j reads them from ``java.text.DateFormatSymbols`` in the language's locale
(``FormattingSwitchHelper.getSimpleDateFormat``); here they come from the
committed CLDR-derived table ``resources/date-names.json``, written by
``scripts/generate_date_names.py``, so that the package takes no dependency for
them (CR-014 Q10). What makes them Word's is docx4j's two rules, kept here and not
in the table:

* an abbreviated month loses a trailing period (``wordMonthAbbreviations``): CLDR's
  German ``Okt.`` is ``Okt``, and a picture's own ``MMM.`` supplies the period;
* AM and PM are upper case, whichever of ``am/pm`` and ``AM/PM`` asked
  (``wordAmPm``; Word 365 showed ``2:05 AM`` for en-AU, whose CLDR names are
  ``am`` and ``pm``).

A language the table doesn't hold has no names: :func:`date_names` answers None,
and the caller shows English and says so (Q10).
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from functools import cache

__all__ = ["DATE_NAMES_RESOURCE", "DateNames", "date_names"]

#: The table, in ``docx4j_py/resources/``.
DATE_NAMES_RESOURCE = "date-names.json"


@dataclass(frozen=True, slots=True)
class DateNames:
    """One language's names, as a date picture writes them."""

    #: January first.
    months: tuple[str, ...]
    #: January first, no trailing period.
    months_short: tuple[str, ...]
    #: Monday first, as ``datetime.weekday()`` counts.
    days: tuple[str, ...]
    #: Monday first.
    days_short: tuple[str, ...]
    #: AM and PM, upper case.
    am_pm: tuple[str, str]


@cache
def _table() -> dict[str, DateNames]:
    """The table, by lower-cased tag, read on first use."""
    from importlib.resources import files

    data = json.loads(files("docx4j_py.resources").joinpath(DATE_NAMES_RESOURCE).read_text("utf-8"))
    return {
        tag.lower(): DateNames(
            months=tuple(names["months"]),
            months_short=tuple(month.rstrip(".") for month in names["months_short"]),
            days=tuple(names["days"]),
            days_short=tuple(names["days_short"]),
            am_pm=(names["am_pm"][0].upper(), names["am_pm"][1].upper()),
        )
        for tag, names in data["locales"].items()
    }


def date_names(lang: str | None) -> DateNames | None:
    """The names of a ``w:lid`` or ``w:lang`` tag: the tag's own, else its language's, else None."""
    tag = (lang or "").strip().replace("_", "-").lower()
    table = _table()
    return table.get(tag) or table.get(tag.split("-", 1)[0])
