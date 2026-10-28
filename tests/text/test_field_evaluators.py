"""The field evaluators, their context and the text model's set. CR-025 Phase B.

docx4j's ``FieldEvaluatorsTest`` at ``c3e9920e5``, the cases that apply (its
updater's are ``FieldUpdater``'s, which this repository has no port of): the
property fields, a pass's sequences and references, the date fields at a
clock, a caller's own evaluator, the core-property dates, ``REVNUM`` and
``INFO``, ``SECTION``, ``NOTEREF`` and the user. Then ``PackageDateTimeZoneTest``'s
field cases: the date fields in the package's zone and the per-call
override. Then this port's decisions: the forms Word had not been measured on
(D9; since Phase B2, Word's reading of them, docx4j's D13: UTC), ``INFO``'s
reach (D12) and what ``fields: update`` keeps cached (D14).
"""

from __future__ import annotations

import datetime
import zoneinfo

import pytest
from mini_docx import MiniDocx

from docx4j_py.model.fields import (
    FieldContext,
    FieldUser,
    evaluate,
    parse_w3cdtf,
    register,
)
from docx4j_py.model.fields import evaluators as registry
from docx4j_py.model.text import TextOptionError, extract, plain_text
from docx4j_py.model.text.field_values import COMPUTED

UTC = datetime.UTC
CORE_TYPE = "application/vnd.openxmlformats-package.core-properties+xml"
CORE_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
NS_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _core(created: str | None = None, modified: str | None = None, revision: str | None = None) -> str:
    dates = "".join(
        f'<dcterms:{name} xsi:type="dcterms:W3CDTF">{value}</dcterms:{name}>'
        for name, value in (("created", created), ("modified", modified))
        if value is not None
    )
    return (
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:title>The Title</dc:title><dc:creator>Ann Author</dc:creator><cp:keywords>alpha, beta</cp:keywords>"
        f"{f'<cp:revision>{revision}</cp:revision>' if revision else ''}{dates}</cp:coreProperties>"
    )


def pkg(body: str = "<w:p/>", **dates: str | None):
    """docx4j's ``FieldEvaluatorsTest.pkg``: title, creator, keywords, a custom Client = Acme, a docVar v1 = one."""
    return (
        MiniDocx.body(body)
        .part(
            "settings.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
            "rIdSettings",
            f'<w:settings {NS_W}><w:docVars><w:docVar w:name="v1" w:val="one"/></w:docVars></w:settings>',
        )
        .package_part("docProps/core.xml", CORE_TYPE, CORE_REL, "rIdCore", _core(**dates))
        .package_part(
            "docProps/custom.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
            "rIdCustom",
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Client"><vt:lpwstr>Acme</vt:lpwstr>'
            "</property></Properties>",
        )
        .load()
    )


def dated_pkg(body: str = "<w:p/>"):
    """docx4j's ``datedPkg``: revision 7, created 2024-02-03T04:05:06Z, modified 2025-06-07T08:09:10Z."""
    return pkg(body, created="2024-02-03T04:05:06Z", modified="2025-06-07T08:09:10Z", revision="7")


# ---------------------------------------------------------------------------
# FieldEvaluatorsTest
# ---------------------------------------------------------------------------


def test_the_property_fields():
    ctx = FieldContext(pkg())
    assert evaluate(" DOCPROPERTY Client ", ctx) == "Acme"
    assert evaluate(' DOCPROPERTY "Client" \\* Upper ', ctx) == "ACME", "the switches are applied"
    assert evaluate(" TITLE ", ctx) == "The Title"
    assert evaluate(" AUTHOR ", ctx) == "Ann Author"
    assert evaluate(' AUTHOR "Jo" ', ctx) == "Jo", "an argument is what Word sets the property from, and shows"
    assert evaluate(" KEYWORDS ", ctx) == "alpha, beta"
    assert evaluate(" DOCVARIABLE v1 ", ctx) == "one"
    assert evaluate(' QUOTE "hi there" ', ctx) == "hi there"
    assert evaluate(" DOCPROPERTY nope ", ctx) is None, "a missing property keeps the cached result"
    assert evaluate(" DOCPROPERTY ", ctx) is None, "no argument"
    assert evaluate(" FILENAME ", ctx) is None, "not evaluated here"
    assert evaluate(" MERGEFIELD x ", ctx) is None, "not evaluated here"
    assert evaluate("", ctx) is None
    assert evaluate(None, ctx) is None


def test_sequences_and_references_are_the_passs():
    ctx = FieldContext(pkg())
    assert evaluate(" SEQ Figure ", ctx) == "1"
    assert evaluate(" SEQ Figure \\* ARABIC ", ctx) == "2"
    assert evaluate(" SEQ Figure \\c ", ctx) == "2", "\\c repeats"
    assert evaluate(" SEQ Table \\r 5 ", ctx) == "5", "\\r resets"
    assert evaluate(" SEQ Table ", ctx) == "6"
    assert evaluate(" SEQ Figure \\h ", ctx) == "", "\\h hides, and counts"
    assert evaluate(" SEQ Figure ", ctx) == "4"
    assert evaluate(" SEQ figure ", ctx) == "5", "the identifier's case does not matter"
    assert evaluate(" SEQ ", ctx) is None
    assert evaluate(" SEQ X \\r abc ", ctx) is None, "a reset that is not a number"

    ctx.bookmarks["bm1"] = "Bookmarked text"
    assert evaluate(" REF bm1 ", ctx) == "Bookmarked text"
    assert evaluate(" REF BM1 \\* Upper ", ctx) == "BOOKMARKED TEXT"
    assert evaluate(" REF bm1 \\n ", ctx) is None, "a REF needing numbering or layout keeps its cached result"
    assert evaluate(" REF nope ", ctx) is None
    another = FieldContext(ctx.package)
    assert evaluate(" SEQ Figure ", another) == "1", "a new context is a new pass"


def test_the_date_fields_read_the_clock(zone):
    zone("Australia/Sydney")
    fixed = datetime.datetime(2026, 3, 4, 5, 6, 7).astimezone()  # docx4j's Calendar.getInstance(), the JVM's zone
    ctx = FieldContext(pkg(), clock=lambda: fixed)
    assert evaluate(' DATE \\@ "yyyy-MM-dd" ', ctx) == "2026-03-04"
    assert evaluate(' TIME \\@ "HH:mm" ', ctx) == "05:06"
    assert evaluate(' DATE \\@ "d MMMM yyyy" ', FieldContext(pkg(), clock=lambda: fixed, language="en-US")) == (
        "4 March 2026"
    )
    # PRINTDATE is cp:lastPrinted, not the clock (docx4j's D14, CR-025 section 13.10): never printed, cached
    assert evaluate(' PRINTDATE \\@ "d MMMM yyyy" ', ctx) is None
    # docx4j's fr-FR case ("4 mars 2026") has no counterpart: the names are English (CR-025 D7)
    assert evaluate(' DATE \\@ "d MMMM yyyy" ', FieldContext(pkg(), clock=lambda: fixed, language="fr-FR")) == (
        "4 March 2026"
    )
    assert registry.formats_itself("DATE")
    assert not registry.formats_itself("DOCPROPERTY")
    assert evaluate(" DATE ", ctx) == "3/4/26", "no picture: Java's short date"
    assert evaluate(" TIME ", ctx) == "5:06 AM", "no picture: Java's short time"


def test_a_callers_own_evaluator():
    assert "DATE" in registry.names()
    assert registry.has("docproperty")
    assert not registry.has("TESTUSER")
    register("TESTUSER", lambda instruction, context: context.user.name if context.user else None)
    try:
        assert registry.has("testuser")
        ctx = FieldContext(pkg())
        assert evaluate(" TESTUSER ", ctx) is None, "the user is absent by default: the field keeps its cached result"
        ctx.user = FieldUser("Jo", "J")
        assert evaluate(" TESTUSER \\* Upper ", ctx) == "JO"
    finally:
        registry._REGISTRY.pop("TESTUSER", None)
    assert not registry.has("TESTUSER")


def test_the_core_property_dates_revision_and_info(zone):
    zone("Australia/Sydney")
    ctx = FieldContext(dated_pkg())
    created = parse_w3cdtf("2024-02-03T04:05:06Z")
    assert created == datetime.datetime(2024, 2, 3, 4, 5, 6, tzinfo=UTC)
    assert created.timestamp() * 1000 == 1706933106000
    assert evaluate(' CREATEDATE \\@ "yyyy-MM-dd" ', ctx) == "2024-02-03"
    assert evaluate(' SAVEDATE \\@ "d MMMM yyyy" ', ctx) == "7 June 2025"
    assert evaluate(" CREATEDATE ", ctx) == "2/3/2024 3:05:00 PM", "Word's default, measured (CR-067 section 11.3)"
    assert evaluate(" REVNUM ", ctx) == "7"
    assert evaluate(" INFO DOCPROPERTY Client ", ctx) == "Acme"
    assert evaluate(" INFO Title ", ctx) == "The Title"
    assert evaluate(" INFO Title \\* Upper ", ctx) == "THE TITLE"
    assert evaluate(" INFO RevNum ", ctx) == "7"
    assert evaluate(" INFO ", ctx) is None, "no name"
    assert evaluate(" INFO Nope ", ctx) is None, "not a field's name"
    assert evaluate(" INFO INFO Title ", ctx) is None, "INFO INFO is not a loop"
    assert parse_w3cdtf("not a date") is None
    assert parse_w3cdtf("2024-02-03T04:05:06+00:00") == created


def test_section_note_reference_and_the_user():
    body = (
        f'<w:p {NS_W}><w:r><w:t xml:space="preserve">A </w:t></w:r><w:bookmarkStart w:id="1" w:name="bm1"/>'
        '<w:r><w:t>Bookmarked</w:t></w:r><w:r><w:t xml:space="preserve"> text</w:t></w:r><w:bookmarkEnd w:id="1"/>'
        '<w:bookmarkStart w:id="2" w:name="bmNote"/><w:r><w:footnoteReference w:id="1"/></w:r><w:bookmarkEnd w:id="2"/>'
        '<w:bookmarkStart w:id="3" w:name="bmNote2"/><w:r><w:footnoteReference w:id="2"/></w:r><w:bookmarkEnd w:id="3"/>'
        '<w:bookmarkStart w:id="4" w:name="bmEnd"/><w:r><w:endnoteReference w:id="1"/></w:r><w:bookmarkEnd w:id="4"/></w:p>'
        f'<w:p {NS_W}><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>'
        f'<w:p {NS_W}><w:r><w:t>in section 2</w:t></w:r></w:p>'
    )
    ctx = FieldContext(pkg(body)).read_bookmarks()
    assert evaluate(" REF bm1 ", ctx) == "Bookmarked text"
    assert evaluate(" NOTEREF bmNote ", ctx) == "1"
    assert evaluate(" NOTEREF bmNote2 \\f ", ctx) == "2"
    assert evaluate(" NOTEREF bmEnd ", ctx) == "1", "endnotes count on their own"
    assert evaluate(" NOTEREF nope ", ctx) is None
    assert evaluate(" SECTION ", ctx) is None, "the consumer has not said which section"
    ctx.section = 2
    assert evaluate(" SECTION ", ctx) == "2"
    assert evaluate(" USERNAME ", ctx) is None
    ctx.user = FieldUser("Jo", "JB", "1 Street")
    assert evaluate(" USERNAME ", ctx) == "Jo"
    assert evaluate(" USERINITIALS ", ctx) == "JB"
    assert evaluate(" USERADDRESS ", ctx) == "1 Street"


def test_a_custom_mark_a_separator_and_a_deletion_are_not_counted():
    """docx4j's walk: a reference with ``w:customMarkFollows`` is not counted, nor one in a ``w:del``."""
    body = (
        f'<w:p {NS_W}><w:r><w:footnoteReference w:customMarkFollows="1" w:id="1"/><w:t>*</w:t></w:r>'
        '<w:del w:id="9" w:author="A"><w:r><w:footnoteReference w:id="2"/></w:r></w:del>'
        '<w:bookmarkStart w:id="5" w:name="n"/><w:r><w:footnoteReference w:id="3"/></w:r><w:bookmarkEnd w:id="5"/>'
        '<w:bookmarkStart w:id="6" w:name="custom"/><w:r><w:footnoteReference w:customMarkFollows="true" w:id="4"/></w:r>'
        '<w:bookmarkEnd w:id="6"/></w:p>'
    )
    ctx = FieldContext(pkg(body))
    assert evaluate(" NOTEREF n ", ctx) == "1"
    assert evaluate(" NOTEREF custom ", ctx) == "1", "a custom mark holds the count it found"


# ---------------------------------------------------------------------------
# PackageDateTimeZoneTest's field cases
# ---------------------------------------------------------------------------


def test_the_date_fields_show_the_packages_zone(zone):
    zone("Australia/Sydney")
    package = pkg(created="2024-02-03T04:05:06Z")
    noon = datetime.datetime(2024, 2, 3, 12, 0, tzinfo=UTC)
    package.date_time_zone = "Asia/Tokyo"  # UTC+9, no daylight time
    ctx = FieldContext(package, clock=lambda: noon, language="en-US")
    assert evaluate(' CREATEDATE \\@ "HH:mm" ', ctx) == "13:05"
    assert evaluate(" CREATEDATE ", ctx) == "2/3/2024 1:05:00 PM"
    assert evaluate(' DATE \\@ "HH:mm" ', ctx) == "21:00"
    assert evaluate(' TIME \\@ "HH:mm" ', ctx) == "21:00"
    ctx.zone = "America/New_York"
    assert evaluate(' DATE \\@ "HH:mm" ', ctx) == "07:00", "the override"
    package.date_time_zone = "UTC"
    assert evaluate(' CREATEDATE \\@ "HH:mm" ', FieldContext(package)) == "04:05"


def test_the_contexts_zone_falls_back_to_the_packages_then_the_processs(zone):
    from docx4j_py.model.dates import PROCESS_ZONE

    zone("Australia/Sydney")
    package = pkg()
    ctx = FieldContext(package)
    assert ctx.zone is PROCESS_ZONE
    package.date_time_zone = "America/New_York"
    assert ctx.zone == zoneinfo.ZoneInfo("America/New_York")
    ctx.zone = "Europe/Paris"
    assert ctx.zone == zoneinfo.ZoneInfo("Europe/Paris"), "the per-call override"
    ctx.zone = None
    assert ctx.zone == zoneinfo.ZoneInfo("America/New_York")
    with pytest.raises(ValueError):
        FieldContext(package, zone="not a zone")


# ---------------------------------------------------------------------------
# this port's decisions
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("created", "package_zone", "default", "switched"),
    [
        # the process's zone, Sydney: UTC+10 in June, UTC+11 in February and January
        ("2025-06-07T08:09:10", None, "6/7/2025 6:09:00 PM", "2025-06-07 18:09:00"),
        ("2024-02-03", None, "2/3/2024 11:00:00 AM", "2024-02-03 11:00:00"),
        ("2024", None, "1/1/2024 11:00:00 AM", "2024-01-01 11:00:00"),
        ("2024-02", None, "2/1/2024 11:00:00 AM", "2024-02-01 11:00:00"),
        # the package's zone over the process's: Tokyo (UTC+9), New York (UTC-5 in winter, UTC-4 in summer)
        ("2025-06-07T08:09:10", "Asia/Tokyo", "6/7/2025 5:09:00 PM", "2025-06-07 17:09:00"),
        ("2024-02-03", "Asia/Tokyo", "2/3/2024 9:00:00 AM", "2024-02-03 09:00:00"),
        ("2025-06-07T08:09:10", "America/New_York", "6/7/2025 4:09:00 AM", "2025-06-07 04:09:00"),
        ("2024-02-03", "America/New_York", "2/2/2024 7:00:00 PM", "2024-02-02 19:00:00"),
        ("2024", "America/New_York", "12/31/2023 7:00:00 PM", "2023-12-31 19:00:00"),
    ],
)
def test_a_date_alone_or_a_time_with_no_offset_is_utc(zone, created, package_zone, default, switched):
    """D9 as Phase B2 made it, docx4j's D13: UTC, as Word 365 reads a core property's (its date-forms probe, D01, D02).

    Until Phase B2 these were nominal, the wall clock as written in any zone
    (Q8: "until docx4j's date-forms probe has run; then both ports follow
    Word"). The field's minute and default apply as to any other.
    """
    zone("Australia/Sydney")
    package = pkg(created=created)
    package.date_time_zone = package_zone
    ctx = FieldContext(package)
    assert evaluate(" CREATEDATE ", ctx) == default
    assert evaluate(' CREATEDATE \\@ "yyyy-MM-dd HH:mm:ss" ', ctx) == switched


@pytest.mark.parametrize(
    ("created", "expected"),
    [
        ("2024-02-03T04:05:06+05:30", "2/3/2024 9:35:00 AM"),
        ("2024-02-03T04:05:59.999Z", "2/3/2024 3:05:00 PM"),
        ("2024-02-03T04:05:06.4950271Z", "2/3/2024 3:05:00 PM"),
        ("2024-02-03Z", "2/3/2024 11:00:00 AM"),
    ],
)
def test_an_offset_is_an_instant_shown_in_the_zone_to_the_minute(zone, created, expected):
    zone("Australia/Sydney")
    assert evaluate(" CREATEDATE ", FieldContext(pkg(created=created))) == expected


def test_a_core_date_that_is_no_date_or_absent_keeps_the_cached_result():
    """docx4j's ``CoreDate``: no date, no value; and an argument is not the value, as the document-information fields' is."""
    assert evaluate(" CREATEDATE ", FieldContext(pkg(created="not a date"))) is None
    assert evaluate(" SAVEDATE ", FieldContext(pkg())) is None
    assert evaluate(' CREATEDATE "x" \\@ "yyyy" ', FieldContext(dated_pkg())) == "2024"


def test_an_empty_picture_is_the_short_date(zone):
    zone("UTC")
    assert evaluate(" CREATEDATE \\@ ", FieldContext(dated_pkg())) == "2/3/24"


def test_info_evaluates_only_what_the_consumer_computes():
    """D12: ``INFO Date`` asks for ``DATE``, which the text model keeps cached; docx4j's asked the registry until its §12.3."""
    clock = lambda: datetime.datetime(2026, 3, 4, 5, 6, 7, tzinfo=UTC)
    package = dated_pkg()
    assert evaluate(' INFO Date \\@ "yyyy" ', FieldContext(package, clock=clock)) == "2026", "every name, unrestricted"
    restricted = FieldContext(package, clock=clock, computes=COMPUTED)
    assert evaluate(' INFO Date \\@ "yyyy" ', restricted) is None
    assert evaluate(' INFO PrintDate \\@ "yyyy" ', restricted) is None, "PRINTDATE is cp:lastPrinted (D14), cached here"
    assert evaluate(" INFO Title ", restricted) == "The Title"
    assert evaluate(' INFO CreateDate \\@ "yyyy" ', restricted) == "2024"


def _field(instr: str, cached: str) -> str:
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
        f'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{cached}</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def test_fields_update_keeps_the_clock_fields_and_info_of_them_cached():
    """D14: ``DATE``, ``TIME`` and ``PRINTDATE`` keep their cached result under ``fields: update``, as docx4j's text model's do."""
    body = "".join(
        f"<w:p>{_field(instr, 'cached')}</w:p>"
        for instr in ('DATE \\@ "yyyy"', "TIME", "PRINTDATE", "INFO PrintDate", "INFO Title", "QUOTE")
    )
    text = plain_text(extract(dated_pkg(body), fields="update"))
    assert text.splitlines() == ["cached", "cached", "cached", "cached", "The Title", ""]
    assert set(COMPUTED) & {"DATE", "TIME", "PRINTDATE"} == set()


def test_section_in_the_main_story_and_cached_elsewhere():
    """``SECTION`` is the main story's section, counted at the walk; in a footnote it keeps its cached result."""
    body = (
        f"<w:p>{_field('SECTION', 'old')}</w:p>"
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>'
        f'<w:p>{_field("SECTION", "old")}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
        '<w:p><w:fldSimple w:instr=" SECTION "><w:r><w:t>old</w:t></w:r></w:fldSimple></w:p>'
    )
    package = (
        MiniDocx.body(body)
        .footnotes(f'<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>{_field("SECTION", "inNote")}</w:p></w:footnote>')
        .load()
    )
    assert plain_text(extract(package, fields="update")).splitlines() == ["1", "", "2[1]", "2", "[1]inNote"]


def test_field_user_is_refused_when_it_is_no_user():
    with pytest.raises(TextOptionError) as raised:
        extract(pkg(), fields="update", field_user=42)
    assert raised.value.option == "field_user"
    with pytest.raises(TextOptionError):
        extract(pkg(), field_user={"name": "Jo", "email": "jo@example.com"})


def test_a_read_leaves_every_part_as_it_was(zone):
    """REQ-204 under the new fields: nothing is unmarshalled, so a save is the bytes it was."""
    zone("Australia/Sydney")
    body = "".join(
        f"<w:p>{_field(instr, 'x')}</w:p>"
        for instr in ("CREATEDATE", "INFO Author", "SECTION", "NOTEREF nope", "USERNAME", "QUOTE")
    )
    package = dated_pkg(body)
    before = package.save()
    extract(package, fields="update", field_user="Jo")
    assert package.save() == before
