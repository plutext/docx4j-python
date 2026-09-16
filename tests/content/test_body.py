"""``Body``: the verbs, the sequence protocol, the parts it sits on.

CR-003 Phase B, section 3.2. Every mutation here is followed by a save, a
reload and a read back, which is section 7's rule.
"""

from __future__ import annotations

import pytest
from conftest import reloaded, sample

from docx4j_py.model.content import Body, ContentError, Paragraph, Table
from docx4j_py.wml import el


def test_insert_paragraph_text_and_the_docx4j_aliases(new_package):
    body = new_package.body
    first = body.insert_paragraph("Hello World")

    assert isinstance(first, Paragraph)
    assert first.text == "Hello World"
    assert first.style == "Normal"

    title = body.insert_paragraph("Title", location="Start", style="Heading 1")
    assert title.style_id == "Heading1"
    assert title.style == "Heading 1"
    assert title.style_built_in == "Heading1"
    assert first.style_built_in == "Normal"

    body.add_styled_paragraph_of_text("Heading2", "Sub")
    body.add_paragraph_of_text("Tab\there")

    assert [p.text for p in body.paragraphs] == ["Title", "Hello World", "Sub", "Tab\there"]
    assert body.text == "Title\nHello World\nSub\nTab\there"
    # the tab became a w:tab: three items in the run
    assert len(body.paragraphs[3].element.content[0].content) == 3

    # the parents are linked on what was inserted
    document_body = new_package.main_document_part.contents.body
    assert first.element.parent is document_body
    assert first.element.content[0].parent is first.element

    back = reloaded(new_package)
    assert [(p.text, p.style_id) for p in back.body.paragraphs] == [
        ("Title", "Heading1"),
        ("Hello World", "Normal"),
        ("Sub", "Heading2"),
        ("Tab\there", "Normal"),
    ]


def test_the_body_is_a_sequence_of_its_blocks(new_package):
    body = new_package.body
    body.insert_paragraph("one")
    body.insert_paragraph("two")
    body.insert_xml("<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>")

    assert len(body) == 3
    assert isinstance(body[0], Paragraph)
    assert body[0].text == "one"
    # CR-003 Phase C: a w:tbl is a Table view; Block is what is left for the
    # block-level elements that have no view of their own (a w:customXml).
    assert isinstance(body[2], Table)
    assert body[2].name == "w:tbl"
    assert [type(block).__name__ for block in body] == ["Paragraph", "Paragraph", "Table"]
    assert [b.text for b in body[0:2]] == ["one", "two"]

    # `in` is a text search, as CR-003 section 3.1 asks
    assert "two" in body
    assert "three" not in body
    assert body[0] in body


def test_insert_text_and_insert_break(new_package):
    body = new_package.body

    span = body.insert_text("first")
    assert span.text == "first"
    assert body.text == "first"

    body.insert_text(" more", location="End")
    assert body.text == "first more"
    body.insert_text("start ", location="Start")
    assert body.text == "start first more"

    body.insert_break("Page")
    assert body.paragraphs[-1].text == "\n"
    assert len(body.paragraphs) == 2

    body.insert_text("everything", location="Replace")
    assert body.text == "everything"
    assert len(body) == 1

    assert reloaded(new_package).body.text == "everything"


def test_clear_and_get_range(new_package):
    body = new_package.body
    body.insert_paragraph("one")
    body.insert_paragraph("two")

    assert body.get_range().text == "one"
    assert body.get_range("End").text == ""
    body.clear()
    assert len(body) == 0
    assert body.text == ""


def test_get_text_takes_a_budget(new_package):
    body = new_package.body
    for index in range(20):
        body.insert_paragraph(f"paragraph {index}")

    assert len(body.get_text(max_chars=30)) == 30
    assert body.get_text(max_chars=30) == body.text[:30]
    assert body.get_text() == body.text


def test_paragraphs_descend_into_tables_and_controls(new_package):
    body = new_package.body
    body.insert_xml(
        "<w:tbl><w:tblPr/><w:tblGrid/>"
        "<w:tr><w:tc><w:p><w:r><w:t>cell one</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>cell two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        "<w:sdt><w:sdtPr><w:tag w:val='t'/></w:sdtPr>"
        "<w:sdtContent><w:p><w:r><w:t>in a control</w:t></w:r></w:p></w:sdtContent></w:sdt>"
    )

    assert [p.text for p in body.paragraphs] == ["cell one", "cell two", "in a control"]
    assert len(body) == 2, "the table and the control are the body's own blocks"
    assert [b.name for b in body.iter_blocks()] == ["w:tbl", "w:sdt"]


def test_search_and_replace_over_the_whole_body(new_package):
    body = new_package.body
    body.insert_paragraph("The quick brown fox")
    body.insert_paragraph("jumps over the lazy dog")

    assert [hit.text for hit in body.search("the")] == ["The", "the"]
    assert [hit.text for hit in body.search("the", match_case=True)] == ["the"]
    assert len(body.search("o", limit=2)) == 2
    assert body.replace_text("fox", "cat") == 1
    assert body.paragraphs[0].text == "The quick brown cat"


def test_body_over_headers_footers_and_the_notes_parts():
    package = sample("Headers.docx")
    main = package.main_document_part

    headers = package.header_parts()
    assert headers, "the sample has headers"
    header_body = headers[0].body
    assert isinstance(header_body, Body)
    assert header_body.prefix.startswith("header:rId")
    assert header_body.text

    # there is at most one of each of these, so they carry no relationship id
    # (CR-003 Phase D, section 3.4's "footnotes/1/0" and "comments/2/0")
    assert main.footnotes_part.body.prefix == "footnotes"
    assert main.endnotes_part.body.prefix == "endnotes"
    # a notes part keeps its notes in w:footnote, not w:content, and a Body
    # finds the list all the same
    assert len(main.footnotes_part.body) == 2


def test_a_part_with_no_block_content_has_no_body():
    package = sample("sample-docx.docx")
    with pytest.raises(ContentError) as raised:
        _ = package.style_definitions_part.body
    assert raised.value.code == "body.no_content"
    assert "header, footer" in raised.value.hint


def test_body_element_is_still_the_typed_w_body(new_package):
    main = new_package.main_document_part
    assert main.body_element is main.contents.body
    assert new_package.body.container is main.contents.body
    # the tree stays reachable: appending to content still works
    new_package.body.content.append(el.p(content=[el.r(content=[el.t("tree")])]))
    assert new_package.body.text == "tree"


def test_sub_bodies_carry_the_prefix(new_package):
    body = new_package.body
    body.insert_xml(
        "<w:tbl><w:tblPr/><w:tblGrid/>"
        "<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    )
    table = body[0].element
    cell = table.content[0].content[0]
    cell_body = body.sub(cell, "body/0/0/0")

    assert cell_body.part is body.part
    assert cell_body.package is body.package
    assert cell_body.prefix == "body/0/0/0"
    assert cell_body.text == "cell"
    cell_body.insert_paragraph("added")
    assert [p.text for p in body.paragraphs] == ["cell", "added"]
