"""A ``w:date`` or ``w:fullDate`` written as Word does not write one: read as Word reads it, saved as Word saves it. CR-025 Phase B2.

docx4j CR-067 section 12.2, its date-forms probe (Word 365 in Sydney,
2026-10-28; the replay is ``tests/text/test_date_forms_probe.py``): a
revision's ``w:date`` is the local clock in every form, a date alone local
midnight (R01); a date control's ``w:fullDate`` the nominal day (B01 to B03);
and Word's save writes each as the nominal clock to the minute labelled ``Z``.

* Item 2, the readers: a date alone, which the exact parse keeps as its text,
  is local midnight in the package's zone for every revision element's
  ``w:date`` and a comment's (``TrackedChange.date``, ``tracking.date_of``,
  ``Comment.creation_date``), and midnight for a date control's
  (``DatePicker.full_date``); ``dateUtc`` still wins. The other forms read as
  they did.
* Item 3, the save's normal form (the user: "item 3 as recommended"): a date
  alone and a date-time with no offset are written as Word's save writes them;
  a value with a ``Z`` or an offset is left exactly as read. The attributes,
  enumerated from the model: :data:`~docx4j_py.openpackaging.normalise.DATE_ATTRIBUTES`.
"""

from __future__ import annotations

import dataclasses
import datetime
import inspect
import io
import re
import zipfile
import zoneinfo
from importlib import import_module

import canon
import pytest
from lxml import etree

from docx4j_py import load
from docx4j_py.model.content.tracking import date_of as content_date_of
from docx4j_py.model.revisions import tracked_changes
from docx4j_py.model.revisions.change import date_of, utc_date, word_date
from docx4j_py.model.values import date_time, date_time_normal_form
from docx4j_py.openpackaging import normalise
from docx4j_py.openpackaging.normalise import (
    DATE_ATTRIBUTES,
    normalise_tree,
    rendered_rows,
)

pytestmark = pytest.mark.filterwarnings("ignore::docx4j_xsdata.exceptions.ConverterWarning")

UTC = datetime.UTC
SYDNEY = zoneinfo.ZoneInfo("Australia/Sydney")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS = f'xmlns:w="{W}" xmlns:w14="{W14}"'

#: 2024-02-03 00:00 in Sydney (AEDT, UTC+11).
SYDNEY_MIDNIGHT = datetime.datetime(2024, 2, 2, 13, 0, tzinfo=UTC)

_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/comments.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'
)
_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/></Relationships>'
)
_DOCUMENT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
    'Target="comments.xml"/></Relationships>'
)


def every_kind(date: str) -> bytes:
    """A ``.docx`` holding every element of :data:`DATE_ATTRIBUTES` but ``w14``'s two, each dated `date`.

    ``w14:customXmlConflictInsRangeStart`` and ``...DelRangeStart`` are root
    elements in the model and in no content model, so a typed document has
    nowhere to hold them; :func:`test_w14s_conflict_range_starts_are_normalised_by_name`
    covers them over a render.
    """
    a = f'w:author="A" w:date="{date}"'
    body = (
        f'<w:p><w:ins w:id="1" {a}><w:r><w:t>ins</w:t></w:r></w:ins>'
        f'<w:del w:id="2" {a}><w:r><w:delText>del</w:delText></w:r></w:del>'
        f'<w:r><w:rPr><w:b/><w:rPrChange w:id="3" {a}><w:rPr/></w:rPrChange></w:rPr><w:t>bold</w:t></w:r></w:p>'
        f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/><w:numberingChange w:id="4" {a} w:original=""/>'
        f'</w:numPr><w:jc w:val="center"/><w:rPr><w:ins w:id="5" {a}/><w:rPrChange w:id="6" {a}><w:rPr/></w:rPrChange>'
        f'</w:rPr><w:pPrChange w:id="7" {a}><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>mark</w:t></w:r></w:p>'
        f'<w:p><w:moveFromRangeStart w:id="8" {a} w:name="m1"/><w:moveFrom w:id="9" {a}><w:r><w:t>moved</w:t></w:r>'
        '</w:moveFrom><w:moveFromRangeEnd w:id="8"/></w:p>'
        f'<w:p><w:moveToRangeStart w:id="10" {a} w:name="m1"/><w:moveTo w:id="11" {a}><w:r><w:t>moved</w:t></w:r>'
        '</w:moveTo><w:moveToRangeEnd w:id="10"/></w:p>'
        f'<w:customXmlInsRangeStart w:id="12" {a}/><w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>'
        '<w:p><w:r><w:t>c</w:t></w:r></w:p></w:sdtContent></w:sdt><w:customXmlInsRangeEnd w:id="12"/>'
        f'<w:customXmlDelRangeStart w:id="13" {a}/><w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent>'
        '<w:p><w:r><w:t>d</w:t></w:r></w:p></w:sdtContent></w:sdt><w:customXmlDelRangeEnd w:id="13"/>'
        f'<w:customXmlMoveFromRangeStart w:id="14" {a}/><w:customXmlMoveFromRangeEnd w:id="14"/>'
        f'<w:customXmlMoveToRangeStart w:id="15" {a}/><w:customXmlMoveToRangeEnd w:id="15"/>'
        f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblPrChange w:id="18" {a}><w:tblPr/></w:tblPrChange></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>'
        f'<w:tr><w:tblPrEx><w:jc w:val="center"/><w:tblPrExChange w:id="19" {a}><w:tblPrEx/></w:tblPrExChange></w:tblPrEx>'
        f'<w:trPr><w:cantSplit/><w:ins w:id="20" {a}/><w:trPrChange w:id="21" {a}><w:trPr/></w:trPrChange></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:cellIns w:id="22" {a}/><w:tcPrChange w:id="23" {a}><w:tcPr/>'
        f'</w:tcPrChange></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:cellDel w:id="24" {a}/></w:tcPr><w:p/></w:tc>'
        f'<w:tc><w:tcPr><w:cellMerge w:id="25" {a} w:vMerge="cont"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'
        f'<w:p><w:sdt><w:sdtPr><w:id w:val="3"/><w:date w:fullDate="{date}"><w:dateFormat w:val="yyyy-MM-dd"/>'
        '<w:lid w:val="en-US"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr>'
        '<w:sdtContent><w:r><w:t>2024-02-03</w:t></w:r></w:sdtContent></w:sdt></w:p>'
        '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>commented</w:t></w:r><w:commentRangeEnd w:id="0"/>'
        '<w:r><w:commentReference w:id="0"/></w:r></w:p>'
        f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:sectPrChange w:id="26" {a}><w:sectPr/></w:sectPrChange></w:sectPr>'
    )
    comments = (
        f'<w:comments {NS}><w:comment w:id="0" w:author="A" w:initials="A" w:date="{date}">'
        "<w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>"
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _TYPES)
        archive.writestr("_rels/.rels", _RELS)
        archive.writestr("word/_rels/document.xml.rels", _DOCUMENT_RELS)
        archive.writestr("word/document.xml", f"<w:document {NS}><w:body>{body}</w:body></w:document>")
        archive.writestr("word/comments.xml", comments)
    return buffer.getvalue()


def _loaded(date: str):
    package = load(io.BytesIO(every_kind(date)))
    # read, so that a save re-marshals them
    _ = package.main_document_part.contents, package.main_document_part.comments_part.contents
    assert package.skipped == []
    return package


def _dates(data: bytes) -> dict[str, list[str]]:
    """Every ``w:date`` and ``w:fullDate`` in each part of a saved package, in document order."""
    archive = zipfile.ZipFile(io.BytesIO(data))
    out = {}
    for name in ("word/document.xml", "word/comments.xml"):
        root = etree.fromstring(archive.read(name))
        out[name] = [
            value
            for element in root.iter()
            for value in (element.get(f"{{{W}}}date"), element.get(f"{{{W}}}fullDate"))
            if value is not None
        ]
    return out


# ---------------------------------------------------------------------------
# which attributes: the model's
# ---------------------------------------------------------------------------


def test_the_date_table_is_the_models():
    """:data:`DATE_ATTRIBUTES` is every element whose bound classes declare a ``w``-namespace ``xsd:dateTime`` attribute.

    The model's 25: ``w:date`` on 21 revision elements of ``w`` and two of
    ``w14``, and on ``w:comment``; ``w:fullDate`` on ``w:date``. Every class
    each element name binds types its attribute ``XmlDateTime`` alone, so a
    row found by name is never another attribute's.
    """
    from docx4j_xsdata.models.datatype import XmlDate, XmlDateTime

    from docx4j_py.el_index import EL_MODULES
    from docx4j_py.runtime import context

    bound, _holders, broken = normalise._bound()
    derived: dict[str, set[str]] = {}
    for element, classes in bound.items():
        if not element.startswith("{"):
            continue
        for cls in classes - broken:
            for var in context().build(cls).get_attribute_vars():
                if normalise._namespace(var.qname) == W and {XmlDateTime, XmlDate} & set(var.types):
                    assert tuple(var.types) == (XmlDateTime,), (element, var.qname)
                    derived.setdefault(element, set()).add(var.qname)
    assert {element: {attribute} for element, attribute in DATE_ATTRIBUTES.items()} == derived
    assert len(DATE_ATTRIBUTES) == 25
    assert sorted({attribute for attribute in DATE_ATTRIBUTES.values()}) == [f"{{{W}}}date", f"{{{W}}}fullDate"]
    assert [e for e, a in DATE_ATTRIBUTES.items() if a == f"{{{W}}}fullDate"] == [f"{{{W}}}date"]
    assert sum(element.startswith(f"{{{W14}}}") for element in DATE_ATTRIBUTES) == 2

    # and the w namespace has no other xsd:dateTime attribute, on any class, bound or not
    others = set()
    for module_name in sorted(set(EL_MODULES.values())):
        for cls in vars(import_module(module_name.removesuffix(".el"))).values():
            if inspect.isclass(cls) and dataclasses.is_dataclass(cls) and cls not in broken:
                for var in context().build(cls).get_attribute_vars():
                    if normalise._namespace(var.qname) == W and XmlDateTime in var.types:
                        others.add(var.qname)
    assert others == {f"{{{W}}}date", f"{{{W}}}fullDate"}


def _date_classes() -> list[type]:
    """Every class of the model with a ``w:date`` attribute (42: the revision elements' and ``w:comment``'s)."""
    from docx4j_py.el_index import EL_MODULES
    from docx4j_py.runtime import context

    out = []
    for module_name in sorted(set(EL_MODULES.values())):
        for cls in vars(import_module(module_name.removesuffix(".el"))).values():
            if not (inspect.isclass(cls) and dataclasses.is_dataclass(cls)) or cls in out:
                continue
            try:
                names = {var.qname for var in context().build(cls).get_attribute_vars()}
            except Exception:  # noqa: BLE001 - inkml's MappingType, as normalise._bound skips it
                names = set()
            if f"{{{W}}}date" in names:
                out.append(cls)
    return out


# ---------------------------------------------------------------------------
# item 2: the readers
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("2025-01-02T03:04:05Z", (2025, 1, 2, 3, 4, 5, 0)),
        ("2025-06-07T08:09:10", (2025, 6, 7, 8, 9, 10, None)),
        ("2025-06-07T08:09:10+10:00", (2025, 6, 7, 8, 9, 10, 600)),
        ("2024-02-03", (2024, 2, 3, 0, 0, 0, None)),
        (" 2024-02-03 ", (2024, 2, 3, 0, 0, 0, None)),
        ("2024-02-03Z", (2024, 2, 3, 0, 0, 0, 0)),
        ("2024-02-03+05:30", (2024, 2, 3, 0, 0, 0, 330)),
        ("2025-06-07T08:09", None),
        ("2024-02", None),
        ("abc", None),
        ("", None),
    ],
)
def test_date_time_reads_words_forms(text, expected):
    """A date alone is midnight (R01, B01); a minute with no seconds, a month alone or no date is None."""
    found = date_time(text)
    if expected is None:
        assert found is None
        return
    assert (found.year, found.month, found.day, found.hour, found.minute, found.second, found.offset) == expected


def test_date_time_takes_the_models_value_and_a_datetime():
    from docx4j_xsdata.models.datatype import XmlDateTime

    value = XmlDateTime(2025, 1, 2, 3, 4, 5, 0, 0)
    assert date_time(value) is value
    assert date_time(None) is None
    assert date_time(datetime.datetime(2025, 1, 2, 3, 4, 5, tzinfo=UTC)) == value
    assert date_time(12) is None


def test_every_class_with_a_w_date_reads_a_date_alone_as_local_midnight():
    """Every revision element's ``w:date`` and a comment's: the one reader, :func:`date_of` (the brief's "every revision element")."""
    classes = _date_classes()
    assert len(classes) == 42
    for cls in classes:
        node = cls(date="2024-02-03")
        assert word_date(node.date, SYDNEY) == SYDNEY_MIDNIGHT, cls.__name__
        assert date_of(node, SYDNEY) == SYDNEY_MIDNIGHT, cls.__name__
        if "date_utc" in {f.name for f in dataclasses.fields(cls)}:
            # dateUtc still wins; a date alone there is UTC midnight, as docx4j's dateOf reads one
            assert date_of(cls(date="2024-02-03", date_utc="2024-02-05"), SYDNEY) == datetime.datetime(
                2024, 2, 5, tzinfo=UTC
            ), cls.__name__


def test_a_date_alone_in_every_revision_and_a_comment_is_local_midnight_in_the_packages_zone(zone):
    """R01 for every kind the engine lists, the content API's view, a comment, and the package's zone over the process's."""
    zone("America/New_York")  # the package's zone is the one that counts
    package = _loaded("2024-02-03")
    package.date_time_zone = "Australia/Sydney"
    changes = tracked_changes(package)
    assert {str(change.kind) for change in changes} >= {
        "run_insertion", "run_deletion", "run_properties", "paragraph_properties", "paragraph_mark_properties",
        "paragraph_mark_insertion", "move", "table_properties", "row_insertion", "section_properties",
    }
    assert {change.date for change in changes} == {SYDNEY_MIDNIGHT}
    assert {change.date for change in package.body.get_tracked_changes()} == {SYDNEY_MIDNIGHT}
    [comment] = package.body.get_comments()
    assert comment.creation_date == SYDNEY_MIDNIGHT
    assert content_date_of("2024-02-03", SYDNEY) == SYDNEY_MIDNIGHT
    package.date_time_zone = None  # the process's: New York's midnight (EST, UTC-5)
    assert {change.date for change in tracked_changes(package)} == {datetime.datetime(2024, 2, 3, 5, 0, tzinfo=UTC)}


def test_the_other_forms_read_as_they_did(zone):
    """A time with no offset is the local clock as written (R02); a ``Z`` the local clock, the ``Z`` ignored (R03)."""
    zone("UTC")
    assert word_date("2025-06-07T08:09:10", SYDNEY) == datetime.datetime(2025, 6, 6, 22, 9, 10, tzinfo=UTC)
    assert word_date("2025-01-02T03:04:05Z", SYDNEY) == datetime.datetime(2025, 1, 1, 16, 4, 5, tzinfo=UTC)
    assert word_date("2025-01-02T03:04:05+05:00", SYDNEY) == datetime.datetime(2025, 1, 1, 22, 4, 5, tzinfo=UTC)
    assert word_date("2024-02-03+05:00", SYDNEY) == datetime.datetime(2024, 2, 2, 19, 0, tzinfo=UTC), "its offset"
    assert word_date("2024-02-03Z", SYDNEY) == SYDNEY_MIDNIGHT, "a Z is the local clock"
    assert word_date("2025-06-07T08:09", SYDNEY) is None
    assert utc_date("2025-06-07T08:09:10") == datetime.datetime(2025, 6, 7, 8, 9, 10, tzinfo=UTC)


def test_a_date_controls_date_alone_is_midnight_its_nominal_day():
    """B01: Word's picker on the day; ``DatePicker.full_date`` midnight, naive, as an offset-free value reads."""
    package = _loaded("2024-02-03")
    [control] = [c for c in package.body.content_controls if c.type == "DatePicker"]
    assert control.date_picker_content_control.full_date == datetime.datetime(2024, 2, 3)  # noqa: DTZ001 - nominal
    assert control.date_picker_content_control.to_dict()["full_date"] == "2024-02-03T00:00:00"


# ---------------------------------------------------------------------------
# item 3: the save's normal form
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("2024-02-03", "2024-02-03T00:00:00Z"),
        ("2025-06-07T08:09:10", "2025-06-07T08:09:00Z"),
        ("2025-06-07T08:09:00", "2025-06-07T08:09:00Z"),
        ("2025-06-07T08:09:59.999", "2025-06-07T08:09:00Z"),
        (" 2024-02-03 ", "2024-02-03T00:00:00Z"),
        # left exactly as read: a Z or an offset (the user's decision), or no date
        ("2025-01-02T03:04:05Z", None),
        ("2025-01-02T03:04:00Z", None),
        ("2025-01-02T03:04:05.5Z", None),
        ("2025-01-02T03:04:05+10:00", None),
        ("2025-01-02T03:04:05-00:00", None),
        ("2024-02-03Z", None),
        ("2024-02-03+05:30", None),
        ("2025-06-07T08:09", None),
        ("abc", None),
        ("", None),
    ],
)
def test_the_normal_form(text, expected):
    assert date_time_normal_form(text) == expected


def test_a_save_writes_words_form_for_a_date_alone_in_every_place():
    """Every ``w:date`` and ``w:fullDate`` of the document and the comments: one row each in ``last_save.normalised``."""
    package = _loaded("2024-02-03")
    data = package.save()
    dates = _dates(data)
    assert len(dates["word/document.xml"]) == 25 and len(dates["word/comments.xml"]) == 1
    assert set(dates["word/document.xml"]) == set(dates["word/comments.xml"]) == {"2024-02-03T00:00:00Z"}
    rows = package.last_save.normalised
    assert len(rows) == 26
    assert {(row.before, row.after) for row in rows} == {("2024-02-03", "2024-02-03T00:00:00Z")}
    assert sorted(row.part for row in rows) == ["/word/comments.xml"] + ["/word/document.xml"] * 25
    assert {row.path.rsplit("/", 2)[-2] for row in rows} == {
        "w:" + name
        for name in (
            "ins", "del", "rPrChange", "numberingChange", "pPrChange", "moveFromRangeStart", "moveFrom",
            "moveToRangeStart", "moveTo", "customXmlInsRangeStart", "customXmlDelRangeStart",
            "customXmlMoveFromRangeStart", "customXmlMoveToRangeStart", "tblPrChange", "tblPrExChange",
            "trPrChange", "cellIns", "tcPrChange", "cellDel", "cellMerge", "date", "sectPrChange", "comment",
        )
    }
    assert rows[0].path == "w:document/w:body/w:p[1]/w:ins/@w:date"
    assert any(row.path.endswith("w:sdt/w:sdtPr/w:date/@w:fullDate") for row in rows)


def test_an_offset_free_time_is_written_to_the_minute_with_a_z():
    package = _loaded("2025-06-07T08:09:10")
    dates = _dates(package.save())
    assert set(dates["word/document.xml"]) == set(dates["word/comments.xml"]) == {"2025-06-07T08:09:00Z"}
    assert len(package.last_save.normalised) == 26


@pytest.mark.parametrize("form", ["2025-01-02T03:04:05Z", "2025-01-02T03:04:05+10:00", "2025-01-02T03:04:05.123Z"])
def test_a_value_with_a_z_or_an_offset_is_left_exactly_as_read(form):
    """The user's decision: no truncation of a well-formed value, though Word's save cuts the seconds."""
    package = _loaded(form)
    dates = _dates(package.save())
    assert package.last_save.normalised == []
    assert set(dates["word/document.xml"]) == set(dates["word/comments.xml"]) == {form}


def test_normalise_false_keeps_the_sources_spelling_and_the_tree_keeps_what_was_read():
    package = _loaded("2024-02-03")
    kept = len(package.main_document_part.kept)
    data = package.save(normalise=False)
    assert package.last_save.normalised == []
    assert set(_dates(data)["word/document.xml"]) == {"2024-02-03"}
    package.save()
    assert len(package.last_save.normalised) == 26
    assert len(package.main_document_part.kept) == kept
    first_ins = package.main_document_part.contents.body.content[0].content[0]
    assert first_ins.date == "2024-02-03", "the render is rewritten, never the tree"
    assert 'w:date="2024-02-03"' in package.main_document_part.xml.decode(), "part.xml outside a save is the tree's"
    assert set(_dates(package.save(normalise=False))["word/document.xml"]) == {"2024-02-03"}


def test_the_saved_form_reads_back_as_the_source_read(zone):
    """Word's meaning kept: local midnight before and after; an offset-free time to the minute, which is what Word holds."""
    zone("Australia/Sydney")
    for source, expected in (
        ("2024-02-03", SYDNEY_MIDNIGHT),
        ("2025-06-07T08:09:10", datetime.datetime(2025, 6, 6, 22, 9, tzinfo=UTC)),
    ):
        package = _loaded(source)
        again = load(io.BytesIO(package.save()))
        assert again.main_document_part.kept == [] or all(
            k.qname.split("}")[-1] not in ("date", "fullDate") for k in again.main_document_part.kept
        )
        assert {change.date for change in tracked_changes(again)} == {expected}
        [comment] = again.body.get_comments()
        assert comment.creation_date == expected
        assert {change.date.replace(second=0) for change in tracked_changes(_loaded(source))} == {expected}


def test_canon_writes_the_same_normal_form_into_the_source():
    """``canon.compare(..., normalise=True)``: 25 and 1 ``word-normal-form`` rows, and nothing else moves."""
    source = every_kind("2024-02-03")
    package = _loaded("2024-02-03")
    saved = zipfile.ZipFile(io.BytesIO(package.save()))
    original = zipfile.ZipFile(io.BytesIO(source))
    for name, count in (("word/document.xml", 25), ("word/comments.xml", 1)):
        report = canon.compare(original.read(name), saved.read(name), normalise=True)
        assert report.identical, report.differences
        assert report.counts == {"word-normal-form": count}
        assert not canon.compare(original.read(name), saved.read(name)).identical


def test_a_value_set_in_code_with_no_offset_is_written_in_the_form_too():
    """The rows read the render, as CR-018 Phase B's do: an ``XmlDateTime`` with no offset, set in code."""
    from docx4j_xsdata.models.datatype import XmlDateTime

    package = _loaded("2025-01-02T03:04:05Z")
    first_ins = package.main_document_part.contents.body.content[0].content[0]
    first_ins.date = XmlDateTime(2025, 6, 7, 8, 9, 10)
    dates = _dates(package.save())
    assert dates["word/document.xml"][0] == "2025-06-07T08:09:00Z"
    assert [(row.before, row.after) for row in package.last_save.normalised] == [
        ("2025-06-07T08:09:10", "2025-06-07T08:09:00Z")
    ]


def test_w14s_conflict_range_starts_are_normalised_by_name():
    """The two ``w14`` elements of the table, which a typed document cannot hold, in a render and in a source."""
    xml = (
        f'<w:document {NS}><w:body><w:p>'
        '<w14:customXmlConflictInsRangeStart w:id="1" w:author="A" w:date="2024-02-03"/>'
        '<w14:customXmlConflictDelRangeStart w:id="2" w:author="A" w:date="2025-06-07T08:09:10"/>'
        '<w:ins w:id="3" w:author="A" w:date="2025-01-02T03:04:05Z"/>'
        '<w:customXml w:element="x" w:date="2024-02-03"/>'  # a w:date where the model has none: not a row
        "</w:p></w:body></w:document>"
    )
    rows = rendered_rows(etree.fromstring(xml), "/word/document.xml")
    assert [(row.path.rsplit("/", 2)[-2], row.before, row.after) for row in rows] == [
        ("w14:customXmlConflictInsRangeStart", "2024-02-03", "2024-02-03T00:00:00Z"),
        ("w14:customXmlConflictDelRangeStart", "2025-06-07T08:09:10", "2025-06-07T08:09:00Z"),
    ]
    source = etree.fromstring(xml)
    assert [(attribute, before, after) for _path, attribute, before, after in normalise_tree(source)] == [
        (f"{{{W}}}date", "2024-02-03", "2024-02-03T00:00:00Z"),
        (f"{{{W}}}date", "2025-06-07T08:09:10", "2025-06-07T08:09:00Z"),
    ]
    assert re.findall(r'w:date="([^"]*)"', etree.tostring(source).decode()) == [
        "2024-02-03T00:00:00Z", "2025-06-07T08:09:00Z", "2025-01-02T03:04:05Z", "2024-02-03",
    ]
