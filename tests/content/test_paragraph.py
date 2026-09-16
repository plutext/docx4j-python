"""``Paragraph``: the style semantics of section 4, the measurements, the verbs."""

from __future__ import annotations

import pytest
from conftest import reloaded, sample

from docx4j_py.model.content import ContentError, Paragraph, Range, StyleError
from docx4j_py.wml import wml


def test_insert_text_insert_paragraph_and_delete(new_package):
    body = new_package.body
    paragraph = body.insert_paragraph("middle")

    paragraph.insert_text("start ", location="Start")
    paragraph.insert_text(" end", location="End")
    assert paragraph.text == "start middle end"
    assert len(paragraph.runs) == 1, "text extends the run that is there"

    span = paragraph.insert_text("replaced", location="Replace")
    assert isinstance(span, Range)
    assert paragraph.text == "replaced"
    assert span.text == "replaced"

    before = paragraph.insert_paragraph("before", location="Before")
    paragraph.insert_paragraph("after", location="After")
    assert [p.text for p in body.paragraphs] == ["before", "replaced", "after"]
    assert before.index == 0

    paragraph.delete()
    assert [p.text for p in body.paragraphs] == ["before", "after"]
    assert paragraph.index == -1

    assert [p.text for p in reloaded(new_package).body.paragraphs] == ["before", "after"]


def test_insert_paragraph_carries_the_properties_but_not_the_section(new_package):
    paragraph = new_package.body.insert_paragraph("one", style="Heading 1")
    paragraph.alignment = "Centered"
    paragraph.element.p_pr.sect_pr = None

    new = paragraph.insert_paragraph("two")
    assert new.style_id == "Heading1"
    assert new.alignment == "Centered"
    assert new.element.p_pr.sect_pr is None
    assert new.element.p_pr.parent is new.element


def test_text_setter_keeps_the_first_run_formatting(new_package):
    body = new_package.body
    paragraph = body.insert_paragraph("")
    paragraph.element.content.append(wml("<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>"))
    paragraph.element.content.append(wml("<w:r><w:t> plain</w:t></w:r>"))
    assert paragraph.text == "bold plain"

    paragraph.text = "set text"
    assert paragraph.text == "set text"
    assert len(paragraph.runs) == 1
    assert paragraph.runs[0].r_pr.b is not None
    assert paragraph.runs[0].parent is paragraph.element

    assert reloaded(new_package).body.paragraphs[0].text == "set text"


def test_alignment_maps_start_and_end(new_package):
    paragraph = new_package.body.insert_paragraph("x")
    assert paragraph.alignment == "Unknown", "no w:jc is Unknown, not Left"

    paragraph.alignment = "Centered"
    assert paragraph.alignment == "Centered"
    assert paragraph.element.p_pr.jc.val == "center"

    # the strict spellings read like the transitional ones (CR-003 section 3.12)
    paragraph.element.p_pr.jc.val = "start"
    assert paragraph.alignment == "Left"
    paragraph.element.p_pr.jc.val = "end"
    assert paragraph.alignment == "Right"
    paragraph.element.p_pr.jc.val = "distribute"
    assert paragraph.alignment == "Justified"

    paragraph.alignment = "Unknown"
    assert paragraph.element.p_pr.jc is None


def test_indents_and_spacing_are_points(new_package):
    paragraph = new_package.body.insert_paragraph("x")

    paragraph.left_indent = 36
    paragraph.right_indent = 18
    paragraph.first_line_indent = -18
    paragraph.space_before = 3
    paragraph.space_after = 6
    paragraph.line_spacing = 24

    ind = paragraph.element.p_pr.ind
    assert (ind.left, ind.right, ind.hanging, ind.first_line) == (720, 360, 360, None)
    assert paragraph.left_indent == 36
    assert paragraph.first_line_indent == -18
    assert paragraph.element.p_pr.spacing.before == 60
    assert paragraph.element.p_pr.spacing.after == 120
    assert paragraph.space_after == 6
    assert paragraph.element.p_pr.spacing.line == 480
    assert paragraph.element.p_pr.spacing.line_rule == "exact"
    assert paragraph.line_spacing == 24

    paragraph.first_line_indent = 12
    assert (paragraph.element.p_pr.ind.hanging, paragraph.element.p_pr.ind.first_line) == (
        None,
        240,
    )

    back = reloaded(new_package).body.paragraphs[0]
    assert (back.left_indent, back.right_indent, back.space_after) == (36, 18, 6)


def test_the_indents_read_and_write_the_strict_names(new_package):
    paragraph = new_package.body.insert_paragraph("x")
    paragraph.element.p_pr = None
    paragraph.element.p_pr = wml("<w:pPr><w:ind w:start='720' w:end='360'/></w:pPr>")

    assert paragraph.left_indent == 36
    assert paragraph.right_indent == 18

    paragraph.left_indent = 18
    assert paragraph.element.p_pr.ind.start == 360, "the paragraph already used w:start"
    assert paragraph.element.p_pr.ind.left is None


def test_line_spacing_reads_the_auto_rule(new_package):
    paragraph = new_package.body.insert_paragraph("x")
    paragraph.element.p_pr = wml("<w:pPr><w:spacing w:line='276' w:lineRule='auto'/></w:pPr>")
    assert paragraph.line_spacing == pytest.approx(13.8)
    paragraph.element.p_pr = wml("<w:pPr><w:spacing w:line='480' w:lineRule='exact'/></w:pPr>")
    assert paragraph.line_spacing == 24
    paragraph.element.p_pr = None
    assert paragraph.line_spacing == 0


def test_outline_level(new_package):
    paragraph = new_package.body.insert_paragraph("x")
    assert paragraph.outline_level == 10, "body text is 10, as Office JS has it"

    paragraph.outline_level = 2
    assert paragraph.element.p_pr.outline_lvl.val == 1
    assert paragraph.outline_level == 2

    paragraph.outline_level = 10
    assert paragraph.element.p_pr.outline_lvl is None


def test_style_style_built_in_and_style_id(new_package):
    paragraph = new_package.body.insert_paragraph("x")
    assert (paragraph.style, paragraph.style_built_in, paragraph.style_id) == (
        "Normal",
        "Normal",
        "Normal",
    )

    paragraph.style_built_in = "Heading1"
    assert paragraph.style_id == "Heading1"
    assert paragraph.style == "Heading 1"

    paragraph.style = "Heading 2"
    assert paragraph.style_id == "Heading2"
    assert paragraph.style_built_in == "Heading2"

    paragraph.style_id = "TOC1"
    assert paragraph.style_built_in == "Toc1", "Word's id spelling, Office JS's value"
    assert paragraph.style == "TOC 1"

    with pytest.raises(StyleError) as raised:
        paragraph.style_built_in = "Other"
    assert raised.value.code == "style.other"


def test_style_reads_the_styles_part_only_when_it_is_unmarshalled():
    package = sample("sample-docx.docx")
    styles = package.style_definitions_part
    assert not styles.is_unmarshalled

    paragraph = package.body.paragraphs[0]
    _ = paragraph.style
    assert not styles.is_unmarshalled, "a read never unmarshals the styles part"

    styles.contents  # noqa: B018 - unmarshal it deliberately
    named = {p.style_id: p.style for p in package.body.paragraphs}
    assert named.get("Heading1", "Heading 1") == "Heading 1"


def test_a_style_the_document_does_not_define_is_refused_with_suggestions():
    package = sample("sample-docx.docx")
    paragraph = package.body.paragraphs[0]

    with pytest.raises(StyleError) as raised:
        paragraph.style = "Headingg 1"
    message = str(raised.value)
    assert raised.value.code == "style.not_found"
    assert "Headingg 1" in message
    assert "the closest are" in message
    assert message.count("'") >= 4, "five closest names are listed"

    # a built-in id is written as it stands, defined or not
    paragraph.style = "Intense Quote"
    assert paragraph.style_id == "IntenseQuote"


def test_para_id_is_allocated_when_the_document_uses_them():
    package = sample("DrawingML_GraphicData_wps.docx")
    package.id_seed = 1234
    with_id = next(p for p in package.body.paragraphs if p.para_id)

    added = with_id.insert_paragraph("added")
    assert added.para_id is not None
    assert len(added.para_id) == 8
    assert int(added.para_id, 16) not in (0,)
    assert added.para_id != with_id.para_id

    # the generator is seedable, so the same document and calls give the same id
    again = sample("DrawingML_GraphicData_wps.docx")
    again.id_seed = 1234
    other = next(p for p in again.body.paragraphs if p.para_id).insert_paragraph("added")
    assert other.para_id == added.para_id

    # and with no seed it is derived from the ids the document already uses,
    # which is deterministic too (CR-003 section 3.4)
    derived = sample("DrawingML_GraphicData_wps.docx")
    first = next(p for p in derived.body.paragraphs if p.para_id).insert_paragraph("added")
    twice = sample("DrawingML_GraphicData_wps.docx")
    second = next(p for p in twice.body.paragraphs if p.para_id).insert_paragraph("added")
    assert first.para_id == second.para_id


def test_a_loaded_document_without_para_ids_gets_none():
    # CR-003 section 3.4: a new paragraph gets a paraId when the document
    # already uses them. 2010-sample1.docx does not, so nothing is stamped.
    package = sample("2010-sample1.docx")
    paragraph = package.body.insert_paragraph("x")
    assert paragraph.para_id is None
    assert paragraph.insert_paragraph("y").para_id is None


def test_a_created_document_always_gets_para_ids(new_package):
    # ... "and always in a created document", which is Phase D's other half:
    # every paragraph an agent makes has an address that survives an insert.
    paragraph = new_package.body.insert_paragraph("x")
    assert paragraph.para_id
    assert paragraph.address == f"w14:{paragraph.para_id}"
    assert paragraph.insert_paragraph("y").para_id != paragraph.para_id


def test_insert_break_in_and_around_the_paragraph(new_package):
    body = new_package.body
    paragraph = body.insert_paragraph("text")

    paragraph.insert_break("Line", location="End")
    assert paragraph.text == "text\n"

    paragraph.insert_break("Page", location="After")
    assert len(body) == 2
    assert body.paragraphs[1].text == "\n"

    with pytest.raises(ContentError) as raised:
        paragraph.insert_break("Page", location="Replace")
    assert raised.value.code == "location.invalid"


def test_insert_xml_merges_a_single_paragraph_fragment(new_package):
    body = new_package.body
    paragraph = body.insert_paragraph("A")

    out = paragraph.insert_xml(
        "<w:p><w:r><w:t xml:space='preserve'> and B</w:t></w:r></w:p>", location="End"
    )
    assert out == [paragraph]
    assert paragraph.text == "A and B"
    assert len(body) == 1

    out = paragraph.insert_xml("<w:p><w:r><w:t>C</w:t></w:r></w:p>", location="After")
    assert len(body) == 2
    assert [p.text for p in body.paragraphs] == ["A and B", "C"]
    assert out[0].element.parent is body.container

    assert [p.text for p in reloaded(new_package).body.paragraphs] == ["A and B", "C"]


def test_views_compare_and_hash_by_element(new_package):
    body = new_package.body
    paragraph = body.insert_paragraph("x")

    assert body.paragraphs[0] == paragraph
    assert body.paragraphs[0] is not paragraph
    assert len({body.paragraphs[0], paragraph}) == 1
    assert body.insert_paragraph("y") != paragraph
    assert paragraph != "x"


def test_repr_str_and_to_dict(new_package):
    paragraph = new_package.body.insert_paragraph("Chapter 2: The one about a long title")
    assert repr(paragraph) == "<Paragraph body/0 'Chapter 2: The one about a lon…'>"
    assert str(paragraph) == paragraph.text

    paragraph.style_built_in = "Heading1"
    as_dict = paragraph.to_dict()
    assert as_dict["style"] == "Heading 1"
    assert as_dict["style_built_in"] == "Heading1"
    assert as_dict["runs"] == 1
    assert as_dict["index"] == 0


def test_get_xml_and_get_content(new_package):
    paragraph = new_package.body.insert_paragraph("x", style="Heading 1")
    xml = paragraph.get_xml()
    assert xml.startswith("<w:p ")
    assert 'w:val="Heading1"' in xml
    assert paragraph.get_content() is paragraph.element.content


def test_paragraph_views_over_a_real_document():
    package = sample("toc.docx")
    body = package.body
    assert len(body.paragraphs) > 20
    assert all(isinstance(p, Paragraph) for p in body.paragraphs)
    headings = [p for p in body.paragraphs if p.style_built_in.startswith("Heading")]
    assert headings
    assert all(p.outline_level <= 10 for p in body.paragraphs)
