"""CR-014 phase 1b (section 2.9, Q10 as amended 2026-10-28): a bound date formatted as Word formats it.

docx4j CR-058 phase 1 (``d8d8bfb71``), its probe of the twelve-hour clock and offsets
(``9b9d053cb``) and its amendment (``c3e9920e5``). The oracle is docx4j's
``BoundDateFormatTest`` and ``DateFieldLanguageTest.twelveHourClockRunsOneToTwelve``, on
its fixtures, copied with a SHA-1 manifest (``tests/fixtures/opendope-cr058/``): each
test says which of theirs it ports. Word 365 ran both at UTC+11 on 23 October 2026, and
docx4j CR-058 section 8 holds its tables, which the tests here quote.

docx4j's DF7 (no ``w:lid``, no document default) takes the JVM's locale; a Python
process has no counterpart, so here it is English (CR-014 section 12).
"""

# ruff: noqa: DTZ001 - a bound date's fields are a wall clock, naive

from __future__ import annotations

import datetime
import zoneinfo
from pathlib import Path

import pytest
from conftest import reloaded

from docx4j_py import create_package, load
from docx4j_py.model.content import ContentError
from docx4j_py.model.customxml import format_date
from docx4j_py.model.customxml.bindings import date_of, document_language
from docx4j_py.model.dates import PROCESS_ZONE, bound_date_zone
from docx4j_py.model.fields import format_word_date
from docx4j_py.model.fields.date_names import date_names
from docx4j_py.model.fields.formatting import uses_date_names
from docx4j_py.model.java_text import UnsupportedFormat
from docx4j_py.wml import el

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "opendope-cr058"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DATA = "urn:cr014"
XPATH = "/ns0:data[1]/ns0:value[1]"
MAPPINGS = f"xmlns:ns0='{DATA}'"
SYDNEY = "Australia/Sydney"


def bound(fixture: str = "typed-controls-dates.docx", *, prepare=None):
    """The fixture bound, and the pass's result."""
    package = load(FIXTURES / fixture)
    if prepare is not None:
        prepare(package)
    return package, package.custom_xml_parts.apply_bindings()


def control(package, alias: str):
    return next(c for c in package.body.content_controls if c.title == alias)


def shows(package, alias: str, expected: str) -> None:
    found = control(package, alias)
    assert found.text == expected, alias
    assert "FF0000" not in found.get_xml(), f"{alias} is red"


# ---------------------------------------------------------------------------
# BoundDateFormatTest.wordsFormatting: DF1 to DF13
# ---------------------------------------------------------------------------

#: docx4j's expectations and Word's, but DF2 to DF4, where docx4j shows the nominal time (D6).
WORDS_FORMATTING = {
    "DF1": "Thu 22 Oct 2026",  # ddd: the weekday, not the day of the year
    "DF2": "3:05 PM",  # am/pm: upper case, as Word; nominal time (D6)
    "DF3": "12:05 PM",  # h at noon is 12, not 0 (Word: "11:05 PM" at UTC+11)
    "DF4": "15:05",
    "DF5": "22 Oktober 2026",  # w:lid de-DE on any machine
    "DF6": "Thursday, 22 October 2026",
    "DF7": "22 October 2026",  # no w:lid: not the run's fr-FR; docx4j's JVM locale, English here
    "DF8": "not-a-date",  # D3: the stored text, not today's date in red
    "DF9": "2026-13-45",
    "DF11": "2026-10-22",
    "DF12": "Day 22 of October",
    "DF13": "Okt. 2026",  # Word's abbreviation, no CLDR period
}


@pytest.mark.parametrize(("alias", "expected"), WORDS_FORMATTING.items(), ids=list(WORDS_FORMATTING))
def test_bound_dates_are_formatted_as_word_formats_them(alias, expected):
    """``boundDatesAreFormattedAsWordFormatsThem``."""
    package, _result = bound()
    shows(package, alias, expected)


def test_an_empty_stored_date_shows_the_placeholder():
    """DF10; D3 refined: an empty value shows the placeholder, as an empty text binding does."""
    package, _result = bound()
    found = control(package, "DF10")
    assert found.is_showing_placeholder
    assert "2026" not in found.get_xml().split("sdtContent", 1)[1]
    assert "FF0000" not in found.get_xml()


def test_the_formats_no_locale_touches():
    """``localeFreeFormatsAgreeToday``."""
    package, _result = bound()
    shows(package, "DF4", "15:05")
    shows(package, "DF11", "2026-10-22")


def test_every_control_of_the_probe_applies_and_saves():
    package, result = bound()
    assert len(result.applied) == 13
    assert result.skipped == ()
    again = reloaded(package)
    assert control(again, "DF5").text == "22 Oktober 2026"


def test_a_stored_date_sets_w_full_date_and_a_non_date_leaves_it():
    package, _result = bound()
    assert control(package, "DF3").date_picker_content_control.full_date.replace(tzinfo=None) == datetime.datetime(
        2026, 10, 22, 12, 5
    )
    assert control(package, "DF8").date_picker_content_control.full_date.year == 2001


# ---------------------------------------------------------------------------
# BoundDateFormatTest.withoutALidTheDocumentDefaultLanguageIsUsed
# ---------------------------------------------------------------------------


def with_styles(lang: str | None):
    """A package of one date control with no ``w:lid``, whose styles name `lang` as the document default."""
    package = create_package()
    package.body.insert_xml(
        f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="3"/><w:date><w:dateFormat w:val="d MMMM yyyy"/></w:date></w:sdtPr>'
        "<w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>"
    )
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>2026-10-22T00:00:00Z</value></data>')
    package.body.content_controls[0].xml_mapping.set_mapping(XPATH, MAPPINGS, part)
    styles = package.style_definitions_part.contents
    r_pr = styles.doc_defaults.r_pr_default.r_pr
    r_pr.lang = el.lang(val=lang) if lang else None
    data = package.save()
    return load(data)


def test_without_a_lid_the_document_default_language_is_used():
    package = with_styles("de-DE")
    assert document_language(package) == "de-DE"
    result = package.custom_xml_parts.apply_bindings()
    assert package.body.content_controls[0].text == "22 Oktober 2026"
    assert result.applied[0].notes == ()


def test_its_own_lid_still_wins():
    package = with_styles("de-DE")
    package.body.content_controls[0].date_picker_content_control.date_display_locale = "en-AU"
    package.custom_xml_parts.apply_bindings()
    assert package.body.content_controls[0].text == "22 October 2026"


def test_neither_is_english():
    package = with_styles(None)
    assert document_language(package) is None
    package.custom_xml_parts.apply_bindings()
    assert package.body.content_controls[0].text == "22 October 2026"


def test_reading_the_document_default_leaves_the_styles_part_unread():
    package = with_styles("de-DE")
    styles = package.style_definitions_part
    package.custom_xml_parts.apply_bindings()
    assert not styles.is_unmarshalled, "a read leaves it byte for byte"


def test_a_styles_part_that_will_not_read_names_no_language():
    assert document_language(None) is None


# ---------------------------------------------------------------------------
# BoundDateFormatTest.theDateTimeZonePropertyConvertsAsWordDoes: the opt-in, in the package's zone
# ---------------------------------------------------------------------------


def test_the_package_zone_alone_converts_nothing():
    """The amendment: ``zone alone, no conversion``."""
    package, _result = bound(prepare=lambda p: setattr(p, "date_time_zone", SYDNEY))
    shows(package, "DF2", "3:05 PM")


def in_sydney(package) -> None:
    package.date_time_zone = SYDNEY
    package.bound_dates_converted_as_word = True


#: Word's run at UTC+11, 2026-10-23: DF1 and the date-only DF11 unshifted.
CONVERTED = {"DF1": "Thu 22 Oct 2026", "DF2": "2:05 AM", "DF3": "11:05 PM", "DF4": "02:05", "DF11": "2026-10-22"}


@pytest.mark.parametrize(("alias", "expected"), CONVERTED.items(), ids=list(CONVERTED))
def test_the_opt_in_converts_as_word_does(alias, expected):
    package, _result = bound(prepare=in_sydney)
    shows(package, alias, expected)


def test_the_opt_in_is_off_by_default_and_never_process_wide():
    package = create_package()
    assert package.bound_dates_converted_as_word is False
    assert bound_date_zone(package) is None
    package.date_time_zone = SYDNEY
    assert bound_date_zone(package) is None, "the zone alone"
    package.bound_dates_converted_as_word = True
    assert bound_date_zone(package) == zoneinfo.ZoneInfo(SYDNEY)
    package.date_time_zone = None
    assert bound_date_zone(package) is PROCESS_ZONE, "no zone: the effective zone, the process's"
    assert bound_date_zone(None) is None


def test_the_conversion_writes_the_local_time_labelled_z():
    """Word's save: ``w:fullDate`` carries the local time labelled ``Z`` (DF2: ``2026-10-23T02:05:00Z``)."""
    package, _result = bound(prepare=in_sydney)
    full = control(package, "DF2").date_picker_content_control.full_date
    assert (full.day, full.hour, full.minute) == (23, 2, 5)


def test_a_dry_run_starts_with_its_packages_opt_in():
    package = load(FIXTURES / "typed-controls-dates.docx")
    in_sydney(package)
    with package.dry_run() as trial:
        assert trial.bound_dates_converted_as_word is True
        trial.custom_xml_parts.apply_bindings()
        assert control(trial, "DF2").text == "2:05 AM"
        trial.bound_dates_converted_as_word = False
    assert package.bound_dates_converted_as_word is True


# ---------------------------------------------------------------------------
# BoundDateFormatTest.twelveOClockAndOffsetsAsWordShowsThem
# ---------------------------------------------------------------------------

#: docx4j's nominal reading, which agrees with Word for a value with no offset.
TWELVE_NOMINAL = {
    "TW1": "1:05 AM",  # 01:05Z, nominal
    "TW2": "01:05 PM",
    "TW3": "12:05 PM 22 Oct",  # the offset's own clock time
    "TW4": "12:05 PM 22 Oct",  # no offset: as Word
    "TW5": "12:05 PM 22 Oct",
    "TW6": "01:05",
}
#: Word's six, measured.
TWELVE_WORD = {
    "TW1": "12:05 PM",
    "TW2": "12:05 AM",
    "TW3": "12:05 PM 22 Oct",
    "TW4": "12:05 PM 22 Oct",
    "TW5": "4:05 AM 23 Oct",
    "TW6": "12:05",
}


@pytest.mark.parametrize(("alias", "expected"), TWELVE_NOMINAL.items(), ids=list(TWELVE_NOMINAL))
def test_twelve_oclock_and_offsets_nominally(alias, expected):
    package, _result = bound("typed-controls-twelve.docx")
    shows(package, alias, expected)


@pytest.mark.parametrize(("alias", "expected"), TWELVE_WORD.items(), ids=list(TWELVE_WORD))
def test_twelve_oclock_and_offsets_as_word_shows_them(alias, expected):
    package, _result = bound("typed-controls-twelve.docx", prepare=in_sydney)
    shows(package, alias, expected)


# ---------------------------------------------------------------------------
# DateFieldLanguageTest
# ---------------------------------------------------------------------------


def at(hour: int) -> datetime.datetime:
    return datetime.datetime(2026, 10, 22, hour, 5)


def test_the_twelve_hour_clock_runs_one_to_twelve():
    """``twelveHourClockRunsOneToTwelve``."""
    assert format_word_date("h:mm am/pm", at(12), "en-AU") == "12:05 PM"
    assert format_word_date("hh:mm AM/PM", at(12), None) == "12:05 PM"
    assert format_word_date("h:mm AM/PM", at(0), "en-AU") == "12:05 AM"
    assert format_word_date("HH:mm", at(0), "en-AU") == "00:05"
    assert format_word_date("h:mm am/pm", at(15), "en-AU") == "3:05 PM"
    assert format_word_date("ddd d MMM yyyy", at(15), "en-AU") == "Thu 22 Oct 2026"


def test_the_languages_docx4js_tests_name():
    """``DateFieldLanguageTest``'s lines 58 and 70: ``6 Eylül 2026`` and CLDR's ``06 Sept 2026``, not Word's ``Sep``."""
    day = datetime.datetime(2026, 9, 6)
    assert format_word_date("d MMMM yyyy", day, "tr-TR") == "6 Eylül 2026"
    assert format_word_date("dd MMM yyyy", day, "de-DE") == "06 Sept 2026"


# ---------------------------------------------------------------------------
# the table, and a language it lacks
# ---------------------------------------------------------------------------


def test_the_table_is_read_by_tag_then_by_language():
    assert date_names("de-DE").months[9] == "Oktober"
    assert date_names("DE-at").months[9] == "Oktober"
    assert date_names("en-AU").months_short[5] == "June", "en-AU's own"
    assert date_names("en-GB").months_short[5] == "Jun", "the language's"
    assert date_names("en_US").am_pm == ("AM", "PM")
    assert date_names("zz-ZZ") is None
    assert date_names(None) is None


def test_docx4js_two_rules_are_the_tables_reader():
    german = date_names("de")
    assert german.months_short[9] == "Okt", "wordMonthAbbreviations: CLDR's Okt. loses its period"
    assert german.days_short[3] == "Do.", "a weekday keeps it, as docx4j's does"
    assert date_names("en-AU").am_pm == ("AM", "PM"), "wordAmPm: CLDR's am and pm, upper-cased"
    assert date_names("tr").am_pm == ("ÖÖ", "ÖS")


def test_the_table_ships_as_a_resource():
    from importlib.resources import files

    assert files("docx4j_py.resources").joinpath("date-names.json").is_file()


def language_control(lid: str, picture: str = "d MMMM yyyy", value: str = "2026-10-22T00:00:00Z"):
    package = create_package()
    package.body.insert_xml(
        f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="3"/><w:date><w:dateFormat w:val="{picture}"/><w:lid w:val="{lid}"/>'
        "</w:date></w:sdtPr><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>"
    )
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>{value}</value></data>')
    package.body.content_controls[0].xml_mapping.set_mapping(XPATH, MAPPINGS, part)
    return package


def test_a_language_the_table_lacks_is_english_and_the_entry_says_so():
    package = language_control("es-ES")
    result = package.custom_xml_parts.apply_bindings()
    assert package.body.content_controls[0].text == "22 October 2026"
    (entry,) = result.applied
    assert [note.code for note in entry.notes] == ["binding.date_language_missing"]
    assert "es-ES" in entry.notes[0].message
    assert any("es-ES" in warning for warning in package.changes[-1].warnings)


def test_a_picture_of_numbers_needs_no_names_and_says_nothing():
    package = language_control("es-ES", "yyyy-MM-dd")
    (entry,) = package.custom_xml_parts.apply_bindings().applied
    assert entry.notes == ()


@pytest.mark.parametrize(
    ("picture", "named"),
    [
        ("d MMMM yyyy", True),
        ("MMM", True),
        ("ddd", True),
        ("dddd", True),
        ("h:mm am/pm", True),
        ("yyyy-MM-dd", False),
        ("HH:mm", False),
        ("'MMMM' d", False),
    ],
)
def test_what_writes_a_name(picture, named):
    assert uses_date_names(picture) is named


def test_a_picture_this_cant_write_shows_the_stored_value_and_says_so():
    """docx4j's traverser writes ``[missing!]`` for one; nothing here writes that."""
    package = language_control("en-AU", "d x yyyy", "2026-10-22")
    result = package.custom_xml_parts.apply_bindings()
    assert package.body.content_controls[0].text == "2026-10-22"
    assert [note.code for note in result.applied[0].notes] == ["binding.date_picture_unsupported"]


# ---------------------------------------------------------------------------
# reading a stored date (parseStoredDate)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("stored", "expected"),
    [
        ("2012-08-19T00:00:00Z", datetime.datetime(2012, 8, 19)),
        ("2012-08-19Z", datetime.datetime(2012, 8, 19)),
        ("2012-08-19T10:00:00+10:00", datetime.datetime(2012, 8, 19, 10)),
        ("2012-08-19", datetime.datetime(2012, 8, 19)),
        ("2012-08-19T10:00", datetime.datetime(2012, 8, 19, 10)),
        ("2012-08-19T10:00:05.250", datetime.datetime(2012, 8, 19, 10, 0, 5, 250000)),
        ("  2012-08-19  ", datetime.datetime(2012, 8, 19)),
        ("not-a-date", None),
        ("2026-13-45", None),
        ("2026-02-30", None),
        ("2012-08-19 10:00", None),
        ("2012-8-19", None),
        ("", None),
    ],
)
def test_a_stored_date_is_read_nominally(stored, expected):
    assert date_of(stored) == expected


def test_a_date_alone_is_not_shifted_by_the_conversion():
    tokyo = zoneinfo.ZoneInfo("Asia/Tokyo")
    assert date_of("2012-08-19Z", tokyo) == datetime.datetime(2012, 8, 19)
    assert date_of("2012-08-19T10:00:00", tokyo) == datetime.datetime(2012, 8, 19, 10), "no offset: local"
    assert date_of("2012-08-19T10:00:00Z", tokyo) == datetime.datetime(2012, 8, 19, 19)
    assert date_of("2012-08-19T10:00:00-05:00", tokyo) == datetime.datetime(2012, 8, 20, 0)


# ---------------------------------------------------------------------------
# format_date, the public helper
# ---------------------------------------------------------------------------


def test_format_date_is_the_field_codes_picture_language():
    """Section 2.9's ``a5/p5`` and ``A10/P10``."""
    when = at(14)
    assert format_date("h:mm am/pm", when) == "2:05 PM"
    assert format_date("h:mm AM/PM", when) == "2:05 PM"
    assert format_date("d MMMM yyyy", when, "de-DE") == "22 Oktober 2026"


def test_format_date_keeps_what_it_read_before_word_said_otherwise():
    """``tt`` and double quotes: unmeasured, kept (an open point of CR-014's), a quoted ``tt`` left alone."""
    when = at(14)
    assert format_date("h:mm tt", when) == "2:05 PM"
    assert format_date('d "of" MMMM', when) == "22 of October"
    assert format_date("'tt' d", when) == "tt 22"


def test_format_date_refuses_what_it_cant_write():
    with pytest.raises(UnsupportedFormat):
        format_date("d x", at(1))


def test_an_ascii_letter_alone_is_a_pattern_letter():
    assert format_word_date("yyyy年M月d日", at(1), "en") == "2026年10月22日"


def test_a_bad_zone_is_still_refused():
    package = create_package()
    with pytest.raises(ContentError):
        package.date_time_zone = "not/a/zone"


# --- the write back keeps a node whose date didn't change (CR-014 section 13.3, O10) ---


def written_back(alias: str, *, prepare=None, date=None) -> str:
    package, _result = bound(prepare=prepare)
    view = control(package, alias)
    if date is not None:
        view.date_picker_content_control.full_date = date
    package.custom_xml_parts.update_from_content_controls()
    return view.xml_mapping.custom_xml_node.text


@pytest.mark.parametrize("prepare", [None, in_sydney], ids=["nominal", "as Word"])
def test_an_unchanged_date_keeps_its_nodes_z(prepare):
    assert written_back("DF2", prepare=prepare) == "2026-10-22T15:05:00Z"


def test_an_unchanged_date_alone_stays_a_date_alone():
    assert written_back("DF11") == "2026-10-22"


def test_a_changed_date_is_written_in_words_form():
    # DF3: its node is its own, where DF2 shares one with DF4, whose unchanged date would write after it
    assert written_back("DF3", date=datetime.datetime(2027, 1, 2, 3, 4)) == "2027-01-02T03:04:00"  # noqa: DTZ001 - a wall clock
