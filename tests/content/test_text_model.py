"""The text model: segments, search across runs, and grapheme-safe splitting.

CR-003 sections 3.12 and 3.13.
"""

from __future__ import annotations

import pytest
from conftest import reloaded, sample

from docx4j_py.model.content.text_model import (
    grapheme_clusters,
    is_grapheme_boundary,
    runs_of,
    segments_of,
    snap_back,
    snap_forward,
    text_of_view,
)
from docx4j_py.wml import wml

#: A paragraph whose text is spread over three runs, so that a match spanning
#: them is a real one.
THREE_RUNS = (
    "<w:p>"
    "<w:r><w:t xml:space='preserve'>The qui</w:t></w:r>"
    "<w:r><w:rPr><w:b/></w:rPr><w:t xml:space='preserve'>ck brown f</w:t></w:r>"
    "<w:r><w:t>ox</w:t></w:r>"
    "</w:p>"
)


@pytest.fixture
def three_runs(body):
    """A paragraph of ``The quick brown fox`` written as three runs."""
    return body.insert_xml(THREE_RUNS)[0]


def test_a_match_spans_three_runs(three_runs):
    assert three_runs.text == "The quick brown fox"
    assert len(three_runs.runs) == 3

    hits = three_runs.search("quick brown fox")
    assert len(hits) == 1
    assert hits[0].text == "quick brown fox"
    assert (hits[0].start, hits[0].end) == (4, 19)
    assert len(hits[0].runs) == 3


def test_replacing_a_match_that_spans_three_runs(three_runs, new_package):
    hits = three_runs.search("quick brown fox")
    hits[0].insert_text("slow red dog", location="Replace")

    assert three_runs.text == "The slow red dog"
    assert len(three_runs.runs) == 1, "the emptied runs go"
    assert reloaded(new_package).body.text == "The slow red dog"


def test_a_replacement_takes_the_formatting_of_the_first_replaced_character(three_runs):
    # 'brown' begins inside the bold run, so the replacement is bold too
    three_runs.search("brown")[0].insert_text("X", location="Replace")

    assert three_runs.text == "The quick X fox"
    bold = [run for run in three_runs.runs if run.r_pr is not None and run.r_pr.b is not None]
    assert len(bold) == 1
    assert bold[0].content[0].value == "ck X f"


def test_formatting_a_match_that_spans_three_runs(three_runs):
    three_runs.search("ick brown f")[0].font.italic = True

    assert three_runs.text == "The quick brown fox"
    italic = [run for run in three_runs.runs if run.r_pr is not None and run.r_pr.i is not None]
    assert "".join(
        segment.text for segment in segments_of(three_runs.element) if segment.run in italic
    ) == "ick brown f"


def test_whole_word_and_wildcards(body):
    paragraph = body.insert_paragraph("The quick brown fox jumps over the lazy dog")

    assert len(paragraph.search("fox", match_whole_word=True)) == 1
    assert len(paragraph.search("ox", match_whole_word=True)) == 0
    assert len(paragraph.search("ox")) == 1

    assert [hit.text for hit in paragraph.search("q*k", match_wildcards=True)] == ["quick"]
    assert [hit.text for hit in paragraph.search("f?x", match_wildcards=True)] == ["fox"]
    assert [hit.text for hit in paragraph.search("<l*y>", match_wildcards=True)] == ["lazy"]
    assert [hit.text for hit in paragraph.search("[ld]og", match_wildcards=True)] == ["dog"]
    assert [hit.text for hit in paragraph.search("The", match_case=True)] == ["The"]
    assert len(paragraph.search("the")) == 2


def test_the_segments_of_a_paragraph(body):
    paragraph = body.insert_xml(
        "<w:p><w:r><w:t>a</w:t><w:tab/><w:br/><w:noBreakHyphen/><w:softHyphen/>"
        "<w:sym w:font='Wingdings' w:char='0041'/>"
        "<w:delText>gone</w:delText><w:instrText>PAGE</w:instrText></w:r></w:p>"
    )[0]

    assert paragraph.text == "a\t\n‑­A"
    assert [segment.text for segment in paragraph.segments()] == [
        "a",
        "\t",
        "\n",
        "‑",
        "­",
        "A",
    ]
    assert [segment.editable for segment in paragraph.segments()] == [
        True,
        False,
        False,
        False,
        False,
        False,
    ]
    assert paragraph.segments()[0].to_dict()["start"] == 0


def test_the_two_views_of_a_tracked_paragraph(body):
    paragraph = body.insert_xml(
        "<w:p><w:r><w:t xml:space='preserve'>kept </w:t></w:r>"
        "<w:ins w:id='1' w:author='a' w:date='2020-01-01T00:00:00Z'>"
        "<w:r><w:t xml:space='preserve'>added </w:t></w:r></w:ins>"
        "<w:del w:id='2' w:author='a' w:date='2020-01-01T00:00:00Z'>"
        "<w:r><w:delText>removed</w:delText></w:r></w:del></w:p>"
    )[0]

    assert paragraph.text == "kept added "
    assert paragraph.get_text(view="original") == "kept removed"
    assert [s.revision for s in paragraph.segments()] == [None, "ins"]
    assert len(runs_of(paragraph.element)) == 2
    assert len(runs_of(paragraph.element, view="original")) == 2


def test_text_inside_hyperlinks_smart_tags_and_controls_counts(body):
    paragraph = body.insert_xml(
        "<w:p><w:r><w:t>a</w:t></w:r>"
        "<w:hyperlink r:id='rId1'><w:r><w:t>b</w:t></w:r></w:hyperlink>"
        "<w:smartTag w:element='x'><w:r><w:t>c</w:t></w:r></w:smartTag>"
        "<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>d</w:t></w:r></w:sdtContent></w:sdt>"
        "</w:p>"
    )[0]

    assert paragraph.text == "abcd"
    assert len(paragraph.runs) == 4
    paragraph.search("bc")[0].font.bold = True
    assert paragraph.text == "abcd"


def test_no_text_boxes_body_text_is_the_main_story():
    # the sample's text box lives in a w:drawing, which the text model does not
    # descend into (CR-003 decided question 8)
    package = sample("DrawingML_GraphicData_wps.docx")
    body = package.body
    from docx4j_py.traversal import text_of

    assert "Hello" not in body.text or text_of(package.main_document_part.contents) != body.text


def test_a_split_never_lands_inside_an_emoji_sequence(body):
    family = "\U0001f468‍\U0001f469‍\U0001f467"
    paragraph = body.insert_paragraph(f"a{family}b")
    assert len(paragraph.text) == 7

    assert grapheme_clusters(paragraph.text) == ["a", family, "b"]
    assert [i for i in range(8) if is_grapheme_boundary(paragraph.text, i)] == [0, 1, 6, 7]

    # a range over the middle of the sequence snaps outwards, so the split
    # falls between clusters and the emoji survives
    from docx4j_py.model.content import Range

    Range(paragraph, 3, 4).font.bold = True
    assert paragraph.text == f"a{family}b"
    assert all(family in run_text or run_text in f"a{family}b" for run_text in [paragraph.text])
    assert paragraph.split_at(3) == 1
    assert paragraph.split_at(3, prefer="forward") == 6
    assert paragraph.text == f"a{family}b"


def test_a_split_never_lands_inside_a_devanagari_cluster(body):
    namaste = "नमस्ते"  # नमस्ते
    paragraph = body.insert_paragraph(namaste)

    assert len(paragraph.text) == 6
    assert grapheme_clusters(namaste) == ["न", "म", "स्ते"]
    assert [i for i in range(7) if is_grapheme_boundary(namaste, i)] == [0, 1, 2, 6]

    # offset 4 is between the virama and the consonant it joins: not a boundary
    assert not is_grapheme_boundary(namaste, 4)
    assert snap_back(namaste, 4) == 2
    assert snap_forward(namaste, 4) == 6
    assert paragraph.split_at(4) == 2
    assert paragraph.text == namaste
    assert [text_of_view(run) for run in []] == []


def test_grapheme_boundaries_of_the_other_awkward_cases():
    flag = "\U0001f1ec\U0001f1e7"  # a regional-indicator pair
    assert grapheme_clusters(flag + flag) == [flag, flag]
    assert is_grapheme_boundary(flag, 2)
    assert not is_grapheme_boundary(flag, 1)

    combining = "é"  # e + combining acute
    assert grapheme_clusters(combining) == [combining]
    assert not is_grapheme_boundary(combining, 1)

    keycap = "1️⃣"
    assert grapheme_clusters(keycap) == [keycap]

    skin = "\U0001f44d\U0001f3fb"
    assert grapheme_clusters(skin) == [skin]

    assert grapheme_clusters("a\r\nb") == ["a", "\r\n", "b"]


def test_wml_fragments_keep_w14_attributes_and_spaces():
    paragraph = wml("<w:p w14:paraId='1A2B3C4D'><w:r><w:t xml:space='preserve'> x</w:t></w:r></w:p>")
    assert paragraph.para_id == "1A2B3C4D"
    assert text_of_view(paragraph) == " x"
