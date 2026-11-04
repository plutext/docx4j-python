#!/usr/bin/env python
"""Write ``docx4j_py/resources/date-names.json``, the month, weekday and AM/PM names a bound date is shown in.

CR-014 phase 1b (section 2.9, Q10): a bound date is formatted in Word's date
picture language, in the language of ``w:lid`` (docx4j CR-058's
``FormattingSwitchHelper.formatDate``), and the names are CLDR's. docx4j reads
them from the JVM's ``java.text.DateFormatSymbols``; this package takes no new
dependency for them (Q10), so they are committed, written here from CLDR through
Babel, as ``scripts/generate_fonts.py`` writes ``resources/fonts/``. Never
hand-edited: a language is added to :data:`LOCALES` and this run again::

    python3 scripts/generate_date_names.py            # any Python with Babel
    python3 scripts/generate_date_names.py --check    # the table as it would be written, compared

Babel is this script's alone, a tool of the checkout's, and no dependency of the
package, whose ``.venv-fork`` does not hold it.

**Nothing here decides anything.** Each entry is CLDR's *format* names: months,
abbreviated months, weekdays (Monday first, Python's ``weekday()``), abbreviated
weekdays, and AM and PM as CLDR spells them. docx4j's two rules for Word stay in
code (:mod:`docx4j_py.model.fields.date_names`): an abbreviated month loses a
trailing period, and AM and PM are upper-cased. The names were compared with
``java.text.DateFormatSymbols`` on OpenJDK 21 (2026-11-05): the same strings for
every locale here.

:data:`LOCALES` is Q10's set, the languages docx4j's tests and Word's measurements
name: ``en``, ``en-AU`` (Word's en-AU probes; CLDR's ``June``, ``July`` and
``Sept``, ``am`` and ``pm``), ``de`` (``BoundDateFormatTest`` DF5 and DF13), ``fr``
(DF7's run language, which Word does not read, but ``w:lid`` may name), and ``tr``
(``DateFieldLanguageTest``'s "6 Eylül 2026"). A tag is looked up whole and then by
its language; anything else is shown in English, and the result's entry says so.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docx4j_py" / "resources" / "date-names.json"

#: Q10's languages. CLDR's own tags, with a hyphen.
LOCALES = ("en", "en-AU", "de", "fr", "tr")


def entry(tag: str) -> dict[str, list[str]]:
    """CLDR's format names for `tag`."""
    from babel import Locale, dates

    locale = Locale.parse(tag.replace("-", "_"))
    months = dates.get_month_names("wide", "format", locale)
    months_short = dates.get_month_names("abbreviated", "format", locale)
    days = dates.get_day_names("wide", "format", locale)
    days_short = dates.get_day_names("abbreviated", "format", locale)
    periods = dates.get_period_names("abbreviated", "format", locale)
    return {
        "months": [months[i] for i in range(1, 13)],
        "months_short": [months_short[i] for i in range(1, 13)],
        "days": [days[i] for i in range(7)],
        "days_short": [days_short[i] for i in range(7)],
        "am_pm": [periods["am"], periods["pm"]],
    }


def table() -> str:
    """The file's text."""
    import babel
    from babel.core import get_cldr_version

    data = {
        "_source": {
            "generator": "scripts/generate_date_names.py",
            "babel": babel.__version__,
            "cldr": get_cldr_version(),
        },
        "locales": {tag: entry(tag) for tag in LOCALES},
    }
    return json.dumps(data, ensure_ascii=False, indent=1) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", action="store_true", help="compare the committed table instead of writing it")
    args = parser.parse_args()
    try:
        text = table()
    except ImportError:
        print("this needs Babel, which the package does not depend on: a Python holding it", file=sys.stderr)
        return 2
    if args.check:
        same = OUT.exists() and OUT.read_text(encoding="utf-8") == text
        print(f"{OUT.relative_to(ROOT)}: {'as generated' if same else 'differs from what this writes'}")
        return 0 if same else 1
    OUT.write_text(text, encoding="utf-8")
    print(f"{OUT.relative_to(ROOT)}: {len(LOCALES)} locales")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
