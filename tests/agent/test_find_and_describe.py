"""``find()`` with context, and ``describe()`` without unmarshalling anything.

CR-003 section 3.4, Phase D.
"""

from __future__ import annotations

import json

from conftest import WITHOUT_PARA_IDS, part_bytes, sample

from docx4j_py.model.content import Description, Range, SearchHit


def test_find_reports_where_the_matches_are(report):
    hits = report.body.find("lazy")

    assert len(hits) == 2
    assert all(isinstance(hit, SearchHit) for hit in hits)
    first = hits[0]
    assert first.match == "lazy"
    assert first.ordinal == "body/1"
    assert first.address.startswith("w14:")
    assert (first.start, first.end) == (35, 39)
    assert first.before == "The quick brown fox jumps over the "
    assert first.after == " dog."
    assert first.snippet == "The quick brown fox jumps over the lazy dog."


def test_context_and_limit_are_the_budget(report):
    hit = report.body.find("lazy", context=4)[0]

    assert hit.before == "the "
    assert hit.after == " dog"
    assert hit.snippet == "the lazy dog"

    assert len(report.body.find("lazy", limit=1)) == 1


def test_a_hit_converts_back_into_a_range_to_edit(report):
    body = report.body
    hit = body.find("brown fox")[0]

    span = hit.range(body)
    assert isinstance(span, Range)
    assert span.text == "brown fox"
    assert body.range_of(hit) == span

    # the address is what survives a tool call, so a hit taken from JSON works
    revived = SearchHit(**json.loads(json.dumps(hit.to_dict())))
    revived.range(body).insert_text("red cat")
    assert body.paragraphs[1].text.startswith("The quick red cat")


def test_find_takes_the_search_options_of_phase_b(report):
    assert report.body.find("QUICK") == report.body.find("quick")
    assert report.body.find("QUICK", match_case=True) == []
    assert report.body.find("laz", match_whole_word=True) == []
    assert len(report.body.find("la?y", match_wildcards=True)) == 2


def test_find_on_a_paragraph_and_on_a_range(report):
    paragraph = report.body.paragraphs[1]

    assert [hit.match for hit in paragraph.find("o")] == ["o", "o", "o", "o"]
    assert paragraph.get_range("Whole").find("fox")[0].start == 16


def test_to_dict_is_json_ready(report):
    hit = report.body.find("lazy")[0].to_dict()

    assert set(hit) == {
        "address",
        "ordinal",
        "para_id",
        "start",
        "end",
        "match",
        "snippet",
        "before",
        "after",
    }
    assert json.loads(json.dumps(hit)) == hit


def test_describe_names_the_styles_the_page_and_the_parts(report):
    description = report.describe()

    assert isinstance(description, Description)
    names = {style.name for style in description.styles}
    assert "Heading 1" in names, "the display name, as paragraph.style takes it"
    in_use = {style.id for style in description.styles if style.in_use}
    assert {"Heading1", "Heading2"} <= in_use
    assert not any(style.in_use for style in description.styles if style.id == "Heading9")

    assert round(description.page.width_pt) == 595, "A4 in points"
    assert round(description.page.height_pt) == 842
    assert description.page.orientation == "portrait"
    assert description.page.margins["top"] == 72.0

    names = {part.name for part in description.parts}
    assert "/word/document.xml" in names
    assert "/word/styles.xml" in names
    assert description.application == "docx4j-python"
    assert description.created

    assert description.tracking_on is False
    assert description.authors == {"comments": [], "revisions": []}
    assert description.skipped == 0


def test_describe_lists_the_headers_and_the_footers():
    description = sample("toc.docx").describe()

    assert description.footers, "toc.docx has two footers"
    assert all(prefix.startswith("footer:rId") for prefix in description.footers)


def test_describe_finds_the_revision_authors():
    description = sample("sample-docx.docx").describe()

    assert description.authors["revisions"], "the sample has a w:ins and a w:del"


def test_describe_lists_the_custom_xml_parts_with_their_namespaces():
    description = sample("invoice2013.docx").describe()

    assert description.custom_xml
    item_id, namespace = description.custom_xml[0]
    assert item_id.startswith("{")
    assert all(isinstance(value, str) for value in (item_id, namespace))


def test_describe_unmarshals_nothing_and_keeps_every_part_byte_for_byte():
    name = WITHOUT_PARA_IDS
    before = part_bytes(sample(name).save())

    package = sample(name)
    description = package.describe()
    assert description.styles

    assert not package.style_definitions_part.is_unmarshalled
    assert not package.main_document_part.is_unmarshalled, "describe reads bytes, not the model"
    settings = package.document_settings_part
    assert settings is None or not settings.is_unmarshalled

    after = part_bytes(package.save())
    assert before == after


def test_describe_is_json_ready(report):
    text = report.describe().to_json()
    loaded = json.loads(text)

    assert loaded["page"]["orientation"] == "portrait"
    assert loaded["tracking_on"] is False
    assert report.describe().style_names(kind="paragraph", in_use=True)
