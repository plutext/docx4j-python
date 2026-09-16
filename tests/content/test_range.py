"""``Range``: a span within one paragraph, its font, its edits, its XML."""

from __future__ import annotations

import pytest
from conftest import reloaded

from docx4j_py.model.content import ContentError, Range, SpanError


def test_a_span_is_offsets_into_the_accepted_text(new_package):
    paragraph = new_package.body.insert_paragraph("The quick brown fox")
    span = paragraph.search("quick brown")[0]

    assert (span.start, span.end) == (4, 15)
    assert span.text == "quick brown"
    assert str(span) == "quick brown"
    assert len(span) == 11
    assert span.paragraphs == [paragraph]
    assert repr(span) == "<Range 4:15 'quick brown'>"
    assert span.to_dict() == {
        "address": paragraph.address,
        "text": "quick brown",
        "start": 4,
        "end": 15,
    }


def test_insert_text_at_every_location(new_package):
    paragraph = new_package.body.insert_paragraph("The cat sat")
    span = paragraph.search("cat")[0]

    span.insert_text("big ", location="Before")
    assert paragraph.text == "The big cat sat"
    assert span.text == "cat", "the span still covers the same text"

    span.insert_text("!", location="After")
    assert paragraph.text == "The big cat! sat"

    span.insert_text("dog", location="Replace")
    assert paragraph.text == "The big dog! sat"
    assert span.text == "dog"

    span.text = "fox"
    assert paragraph.text == "The big fox! sat"

    with pytest.raises(ContentError) as raised:
        span.insert_text("x", location="Nowhere")
    assert raised.value.code == "location.invalid"

    assert reloaded(new_package).body.text == "The big fox! sat"


def test_delete_and_get_range(new_package):
    paragraph = new_package.body.insert_paragraph("The quick brown fox")
    span = paragraph.search(" brown")[0]
    span.delete()

    assert paragraph.text == "The quick fox"
    assert (span.start, span.end) == (9, 9)
    assert span.text == ""

    whole = paragraph.get_range()
    assert whole.get_range("Start").to_dict()["end"] == 0
    assert whole.get_range("End").start == len(paragraph.text)


def test_font_splits_the_runs_at_the_boundaries(new_package):
    paragraph = new_package.body.insert_paragraph("plain bold plain")
    span = paragraph.search("bold")[0]

    assert len(paragraph.runs) == 1
    span.font.bold = True
    assert len(paragraph.runs) == 3, "split at both ends"
    assert paragraph.text == "plain bold plain"
    assert paragraph.runs[0].r_pr is None, "the other runs are untouched"
    assert paragraph.runs[2].r_pr is None
    assert paragraph.runs[1].r_pr.b is not None
    assert paragraph.runs[1].r_pr.b_cs is not None, "the complex-script twin too"

    assert span.font.bold is True
    assert paragraph.font.bold is False, "the paragraph's font reads the first run"

    back = reloaded(new_package).body.paragraphs[0]
    assert back.text == "plain bold plain"
    assert len(back.runs) == 3


def test_font_on_an_empty_span_writes_nothing(new_package):
    paragraph = new_package.body.insert_paragraph("text")
    empty = paragraph.get_range("Start")
    empty.font.bold = True
    assert paragraph.runs[0].r_pr is None
    assert len(paragraph.runs) == 1


def test_style_properties_delegate_to_the_paragraph(new_package):
    paragraph = new_package.body.insert_paragraph("text", style="Heading 1")
    span = paragraph.get_range()

    assert span.style == "Heading 1"
    assert span.style_id == "Heading1"
    assert span.style_built_in == "Heading1"

    span.style_built_in = "Heading2"
    assert paragraph.style_id == "Heading2"
    span.style_id = "Heading3"
    assert paragraph.style_id == "Heading3"


def test_search_and_replace_within_a_span(new_package):
    paragraph = new_package.body.insert_paragraph("one two one two one")
    span = Range(paragraph, 4, 15)  # 'two one two'

    hits = span.search("one")
    assert [(h.start, h.end) for h in hits] == [(8, 11)]
    assert span.replace_text("two", "TWO") == 2
    assert paragraph.text == "one TWO one TWO one"


def test_runs_of_a_span(new_package):
    paragraph = new_package.body.insert_paragraph("abc def ghi")
    paragraph.search("def")[0].font.italic = True
    assert len(paragraph.runs) == 3

    covering = paragraph.search("c def g")[0]
    assert len(covering.runs) == 3, "a run partly inside counts"
    assert len(paragraph.search("def")[0].runs) == 1


def test_get_xml_is_the_paragraph_trimmed_to_the_span(new_package):
    paragraph = new_package.body.insert_paragraph("The quick brown fox", style="Heading 1")
    span = paragraph.search("quick brown")[0]

    xml = span.get_xml()
    assert xml.startswith("<w:p ")
    assert ">quick brown<" in xml
    assert "Heading1" in xml, "the paragraph's properties come with it"
    assert "The " not in xml
    assert paragraph.text == "The quick brown fox", "the document is untouched"


def test_the_original_view_is_not_offered_on_a_span(new_package):
    paragraph = new_package.body.insert_paragraph("text")
    span = paragraph.get_range()

    assert span.get_text() == "text"
    with pytest.raises(SpanError) as raised:
        span.get_text(view="original")
    assert raised.value.code == "range.no_original_view"
    assert "paragraph" in raised.value.hint


def test_insert_paragraph_from_a_span(new_package):
    body = new_package.body
    paragraph = body.insert_paragraph("one")
    paragraph.get_range().insert_paragraph("two")
    assert [p.text for p in body.paragraphs] == ["one", "two"]


def test_spans_compare_and_hash(new_package):
    paragraph = new_package.body.insert_paragraph("text")
    assert Range(paragraph, 0, 2) == Range(paragraph, 0, 2)
    assert Range(paragraph, 0, 2) != Range(paragraph, 0, 3)
    assert len({Range(paragraph, 0, 2), Range(paragraph, 0, 2)}) == 1
