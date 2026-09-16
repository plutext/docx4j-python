"""Errors an agent can act on: every message says what to do instead.

CR-003 section 3.1 and section 3.4's list --- an address that no longer exists
names the nearest one, an element a container cannot hold names what it takes, a
style that does not exist lists the five closest, a span that crosses a run
holder says where to split.
"""

from __future__ import annotations

import json

import pytest
from conftest import sample

from docx4j_py.model.content import (
    AddressError,
    ContentError,
    InvalidTargetError,
    SpanError,
    StyleError,
)


def test_every_error_is_json_ready(report):
    with pytest.raises(ContentError) as raised:
        report.element_at("body/99")
    payload = raised.value.to_dict()

    assert json.loads(json.dumps(payload)) == payload
    assert set(payload) == {"code", "message", "hint"}


def test_an_address_error_names_the_nearest_surviving_address(report):
    report.body.paragraphs[-1].delete()

    with pytest.raises(AddressError) as raised:
        report.element_at("body/6")
    assert raised.value.code == "address.not_found"
    assert "'body/5'" in raised.value.message
    assert "outline()" in raised.value.hint


def test_an_invalid_target_names_what_the_container_takes(report):
    from docx4j_py.wml import r

    with pytest.raises(InvalidTargetError) as raised:
        report.body.insert_element(r("a run is not a block"))
    assert raised.value.code == "target.invalid"
    assert "w:body cannot hold w:r" in raised.value.message
    assert "w:p" in raised.value.message
    assert "wrap a run in a w:p" in raised.value.hint


def test_a_style_error_lists_the_closest_names(report):
    with pytest.raises(StyleError) as raised:
        report.body.paragraphs[0].style = "Headding 1"
    assert raised.value.code == "style.not_found"
    assert "the closest are" in raised.value.message
    assert "style_id" in raised.value.hint


def test_a_span_that_crosses_a_run_holder_says_where_to_split(new_package):
    body = new_package.body
    body.insert_xml(
        "<w:p><w:r><w:t>see </w:t></w:r>"
        "<w:hyperlink><w:r><w:t>here</w:t></w:r></w:hyperlink>"
        "<w:r><w:t> now</w:t></w:r></w:p>"
    )
    paragraph = body.paragraphs[-1]
    assert paragraph.text == "see here now"

    inside = paragraph.get_range()
    inside.start, inside.end = 4, 8
    assert inside.holder_boundaries() == []
    inside.require_one_holder()  # a span inside the hyperlink is fine

    crossing = paragraph.search("see here")[0]
    assert crossing.holder_boundaries() == [4]
    with pytest.raises(SpanError) as raised:
        crossing.require_one_holder("a comment anchor")
    assert raised.value.code == "range.crosses_holder"
    assert "a comment anchor needs a span inside one run holder" in raised.value.message
    assert "crosses one at 4" in raised.value.message
    assert "split the span at 4" in raised.value.hint


def test_a_range_has_no_original_view():
    package = sample("sample-docx.docx")
    span = package.body.paragraphs[0].get_range()

    with pytest.raises(SpanError) as raised:
        span.get_text(view="original")
    assert raised.value.code == "range.no_original_view"
    assert "on the paragraph" in raised.value.hint


def test_a_session_error_says_how_to_recover():
    from docx4j_py import DocumentSession

    with pytest.raises(ContentError) as raised:
        DocumentSession().get("gone")
    assert raised.value.code == "session.unknown_handle"
