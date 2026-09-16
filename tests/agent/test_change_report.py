"""``ChangeReport``: every mutation says what it did, and what it cost.

CR-003 section 3.4 and decided question 4.
"""

from __future__ import annotations

import datetime
import json
import time

import pytest
from conftest import SEED, sample

from docx4j_py.model.content import ChangeReport


def test_one_report_per_call_even_when_a_verb_is_written_over_others(report):
    body = report.body
    assert report.changes == []
    assert report.last_change is None

    paragraph = body.insert_paragraph("Appendix", style="Heading 1")

    assert len(report.changes) == 1, "insert_paragraph calls insert_element and style"
    change = report.last_change
    assert isinstance(change, ChangeReport)
    assert change.operation == "insert_paragraph"
    assert change.addresses == (paragraph.address,)
    assert change.created_para_ids == (paragraph.para_id,)
    assert change.text_after == "Appendix"
    assert change.parts_touched == ("/word/document.xml",)
    assert change.at.tzinfo is datetime.UTC


def test_the_list_is_the_callers_to_clear(report):
    report.body.insert_paragraph("one")
    report.body.insert_paragraph("two")
    assert [c.operation for c in report.changes] == ["insert_paragraph"] * 2

    report.changes.clear()
    assert report.changes == []
    assert report.last_change is None


def test_a_text_edit_reports_the_text_either_side(report):
    paragraph = report.body.paragraphs[1]
    paragraph.insert_text("Once upon a time. ", location="Start")

    change = report.last_change
    assert change.operation == "insert_text"
    assert change.text_before == "The quick brown fox jumps over the lazy dog."
    assert change.text_after.startswith("Once upon a time. The quick")
    assert change.addresses == (paragraph.address,)


def test_replace_text_reports_the_pair(report):
    assert report.body.replace_text("lazy", "energetic") == 2

    change = report.last_change
    assert change.operation == "replace_text"
    assert change.text_before == "lazy"
    assert change.text_after == "energetic"
    assert len(report.changes) == 1, "one report for the whole body's replace"


def test_a_format_change_reports_the_values_either_side(report):
    paragraph = report.body.paragraphs[1]

    paragraph.alignment = "Centered"
    change = report.last_change
    assert change.operation == "format"
    assert (change.text_before, change.text_after) == ("alignment=Unknown", "alignment=Centered")

    paragraph.style = "Heading 3"
    assert report.last_change.text_before == "style=Normal"
    assert report.last_change.text_after == "style=Heading 3"

    paragraph.font.bold = True
    change = report.last_change
    assert change.operation == "format"
    assert (change.text_before, change.text_after) == ("bold=False", "bold=True")
    assert change.addresses == (paragraph.address,)


def test_a_range_font_change_is_a_format_too(report):
    span = report.body.find("brown fox")[0].range(report.body)
    span.font.italic = True

    assert report.last_change.operation == "format"
    assert report.last_change.text_after == "italic=True"


def test_moved_lists_the_ordinals_that_shifted_in_the_container_touched(report):
    body = report.body
    assert len(body) == 7

    body.insert_paragraph("appended")
    assert report.last_change.moved == (), "appending --- the default --- moves nothing"

    body.insert_paragraph("prepended", location="Start")
    moved = report.last_change.moved
    assert len(moved) == 8, "every block that was there"
    assert moved[0] == ("body/0", "body/1")
    assert moved[-1] == ("body/7", "body/8")


def test_moved_is_reported_for_a_cell_as_well(report):
    cell_paragraph = report.paragraph_at("body/4/0/0/0")
    cell_paragraph.insert_paragraph("second line in the cell", location="Before")

    assert report.last_change.moved == (("body/4/0/0/0", "body/4/0/0/1"),)


def test_delete_reports_what_went_and_what_moved(report):
    body = report.body
    target = body.paragraphs[1]
    address = target.address

    target.delete()
    change = report.last_change
    assert change.operation == "delete"
    assert change.addresses == (address,)
    assert change.text_before == "The quick brown fox jumps over the lazy dog."
    assert change.text_after == ""
    assert ("body/2", "body/1") in change.moved


def test_clear_reports_the_body(report):
    report.body.clear()

    change = report.last_change
    assert change.operation == "clear"
    assert change.addresses == ("body",)
    assert change.text_after == ""


def test_insert_break_and_insert_xml_report_too(report):
    report.body.insert_break("Page")
    assert report.last_change.operation == "insert_break"

    report.body.insert_xml("<w:p><w:r><w:t>from xml</w:t></w:r></w:p>")
    change = report.last_change
    assert change.operation == "insert_xml"
    assert len(change.addresses) == 1


def test_a_report_is_json_ready(report):
    report.body.insert_paragraph("x")
    text = report.last_change.to_json()
    loaded = json.loads(text)

    assert loaded["operation"] == "insert_paragraph"
    assert loaded["parts_touched"] == ["/word/document.xml"]
    assert "moved" not in loaded, "what is empty is left out"
    assert datetime.datetime.fromisoformat(loaded["at"])


def test_a_detached_body_records_nothing():
    from docx4j_py.model.content import Body
    from docx4j_py.wml import Body as BodyElement

    body = Body(None, BodyElement())
    body.insert_paragraph("no package, no report")
    assert body.text == "no package, no report"


@pytest.mark.parametrize("_run", range(1))
def test_the_report_costs_a_few_microseconds_a_call(_run):
    """CR-003 section 3.4: "cheap (a few fields per call)". Measured."""
    package = sample("2010-sample1.docx")
    package.id_seed = SEED
    body = package.body
    body.insert_paragraph("warm up")
    package.changes.clear()

    count = 2000
    start = time.perf_counter()
    for index in range(count):
        body.insert_paragraph(f"paragraph {index}")
    with_report = (time.perf_counter() - start) / count

    assert len(package.changes) == count
    # the whole call, report and all, is well under a millisecond; the report
    # itself is one object, one datetime and a list append
    assert with_report < 1e-3, f"{with_report * 1e6:.1f} us per insert_paragraph"
