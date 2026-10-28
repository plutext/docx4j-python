"""The python-docx facade, member by member. CR-003 section 3.6, Phase J.

Every read against a known value, and every write followed by **save, reload
and read back**, which is the rule CR-003 section 7 sets for every phase: a
setter that writes markup Word or this engine cannot read again is not
implemented, however green the in-memory assertion is.
"""

from __future__ import annotations

import datetime as dt
import io
import zipfile

import pytest
from conftest import SAMPLES, part_bytes

from docx4j_py.docx import (
    WD_ALIGN_PARAGRAPH,
    WD_ALIGN_VERTICAL,
    WD_BREAK,
    WD_COLOR_INDEX,
    WD_LINE_SPACING,
    WD_ORIENT,
    WD_ROW_HEIGHT_RULE,
    WD_SECTION,
    WD_STYLE_TYPE,
    WD_TABLE_ALIGNMENT,
    WD_UNDERLINE,
    Cm,
    Document,
    Emu,
    Inches,
    Length,
    Mm,
    Pt,
    RGBColor,
    Twips,
)

SEED = 20260918


@pytest.fixture
def doc() -> Document:
    """A new document, seeded so that the ids are reproducible."""
    document = Document()
    document.package.id_seed = SEED
    return document


@pytest.fixture(scope="module")
def png() -> bytes:
    """The pangolin from ``samples/Images.docx``."""
    with zipfile.ZipFile(SAMPLES / "Images.docx") as archive:
        name = next(n for n in archive.namelist() if n.lower().endswith(".png"))
        return archive.read(name)


def reopen(document: Document) -> Document:
    """Save to bytes and open again: what Word would see."""
    return Document(io.BytesIO(document.save()))


# ---------------------------------------------------------------------------
# the units
# ---------------------------------------------------------------------------


def test_the_units_are_python_docxs_to_the_emu():
    assert Pt(24) == 304800
    assert Inches(1) == 914400
    assert Cm(2.54) == 914400
    assert Mm(25.4) == 914400
    assert Twips(1440) == 914400
    assert Emu(914400) == 914400


def test_length_converts_the_way_python_docx_documents_it():
    length = Inches(2)
    assert length.inches == 2.0
    assert length.cm == 5.08
    assert length.mm == 50.8
    assert length.pt == 144.0
    assert length.twips == 2880
    assert length.emu == 1828800
    assert Pt(24).pt == 24.0


def test_length_is_an_int_so_arithmetic_works_and_gives_an_int():
    assert Inches(1) + Inches(1) == Inches(2)
    assert isinstance(Inches(1) + Inches(1), int)
    assert not isinstance(Inches(1) + Inches(1), Length)
    assert Emu(Inches(1) + Inches(1)).inches == 2.0
    assert max(Pt(10), Pt(12)) == Pt(12)


def test_rgbcolor_is_a_triple_that_prints_as_six_hex_digits():
    colour = RGBColor(0x3C, 0x2F, 0x80)
    assert tuple(colour) == (0x3C, 0x2F, 0x80)
    assert str(colour) == "3C2F80"
    assert repr(colour) == "RGBColor(0x3c, 0x2f, 0x80)"
    assert RGBColor.from_string("3C2F80") == colour
    assert RGBColor.from_string("#3c2f80") == colour
    with pytest.raises(ValueError, match="0-255"):
        RGBColor(-1, 0, 0)
    with pytest.raises(TypeError, match="integer"):
        RGBColor("3C", 0, 0)  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# Document
# ---------------------------------------------------------------------------


def test_a_new_document_is_create_packages_a4_with_docx4js_styles(doc):
    assert doc.paragraphs == []
    assert doc.tables == []
    assert len(doc.sections) == 1
    section = doc.sections[0]
    assert round(section.page_width.cm, 2) == 21.0
    assert round(section.page_height.cm, 2) == 29.7
    assert round(section.left_margin.cm, 2) == 2.54
    assert "Heading 1" in doc.styles
    assert doc.package.compatibility_mode == 15


def test_document_opens_a_path_a_stream_and_bytes():
    from_path = Document(SAMPLES / "2010-sample1.docx")
    assert from_path.paragraphs[0].text.startswith("My first")
    data = (SAMPLES / "2010-sample1.docx").read_bytes()
    assert Document(io.BytesIO(data)).paragraphs[0].text == from_path.paragraphs[0].text
    assert Document(data).paragraphs[0].text == from_path.paragraphs[0].text


def test_add_paragraph_and_add_run_and_the_text_setter(doc):
    paragraph = doc.add_paragraph("Plain, ")
    run = paragraph.add_run("bold")
    run.bold = True
    assert paragraph.text == "Plain, bold"
    assert [r.text for r in paragraph.runs] == ["Plain, ", "bold"]

    paragraph.text = "replaced"
    assert paragraph.text == "replaced"
    assert len(paragraph.runs) == 1, "the text setter replaces every run"
    assert paragraph.runs[0].bold is None, "run formatting goes with the runs"

    again = reopen(doc)
    assert [p.text for p in again.paragraphs] == ["replaced"]


def test_a_tab_is_a_w_tab_and_a_newline_a_w_br(doc):
    paragraph = doc.add_paragraph("a\tb\nc")
    from docx4j_py.fragments import to_xml

    xml = to_xml(paragraph.element)
    assert "<w:tab/>" in xml
    assert "<w:br/>" in xml
    assert paragraph.text == "a\tb\nc"
    assert reopen(doc).paragraphs[0].text == "a\tb\nc"


def test_a_carriage_return_is_a_line_break_too(doc):
    doc.add_paragraph("a\rb")
    assert reopen(doc).paragraphs[0].text == "a\nb"


def test_add_heading_levels_0_to_9_and_the_definition_that_follows(doc):
    assert doc.add_heading("Title", 0).style.name == "Title"
    for level in range(1, 10):
        paragraph = doc.add_heading(f"H{level}", level)
        assert paragraph.style.name == f"Heading {level}"
    with pytest.raises(ValueError, match="0-9"):
        doc.add_heading("no", 10)

    # create_package() defines Heading 1 to 4; the rest were added on demand
    # (CR-003 section 14.9), so they survive the round trip as real styles
    again = reopen(doc)
    assert [p.style.name for p in again.paragraphs] == ["Title"] + [
        f"Heading {n}" for n in range(1, 10)
    ]
    assert "Heading 9" in again.styles


def test_add_page_break_writes_a_page_break_run(doc):
    paragraph = doc.add_page_break()
    from docx4j_py.fragments import to_xml

    assert 'w:type="page"' in to_xml(paragraph.element)
    assert paragraph.text == "", "python-docx gives a page break no newline"
    assert reopen(doc).paragraphs[0].text == ""


def test_add_break_takes_every_break_type(doc):
    from docx4j_py.fragments import to_xml

    run = doc.add_paragraph().add_run()
    run.add_break()
    run.add_break(WD_BREAK.PAGE)
    run.add_break(WD_BREAK.COLUMN)
    run.add_break(WD_BREAK.LINE_CLEAR_LEFT)
    xml = to_xml(run.element)
    assert "<w:br/>" in xml
    assert 'w:type="page"' in xml
    assert 'w:type="column"' in xml
    assert 'w:type="textWrapping" w:clear="left"' in xml


def test_insert_paragraph_before_and_clear(doc):
    second = doc.add_paragraph("second")
    first = second.insert_paragraph_before("first", style="Heading 1")
    assert [p.text for p in doc.paragraphs] == ["first", "second"]
    assert first.style.name == "Heading 1"
    first.clear()
    assert first.text == ""
    assert first.style.name == "Heading 1", "clear keeps the paragraph formatting"


def test_iter_inner_content_gives_paragraphs_and_tables_in_order(doc):
    doc.add_paragraph("one")
    doc.add_table(1, 1)
    doc.add_paragraph("two")
    kinds = [type(item).__name__ for item in doc.iter_inner_content()]
    assert kinds == ["Paragraph", "Table", "Paragraph"]


def test_paragraphs_are_the_bodys_own_children_not_the_cells(doc):
    doc.add_paragraph("body")
    doc.add_table(1, 1).cell(0, 0).text = "cell"
    assert [p.text for p in doc.paragraphs] == ["body"], "python-docx's rule"
    # the content API descends, and is one attribute away
    assert [p.text for p in doc.part.body.paragraphs] == ["body", "cell"]


# ---------------------------------------------------------------------------
# Font: the tri-state semantics
# ---------------------------------------------------------------------------


def test_a_boolean_is_three_states_and_false_is_written_not_removed(doc):
    from docx4j_py.fragments import to_xml

    paragraph = doc.add_paragraph()
    inherit, on, off = paragraph.add_run("a"), paragraph.add_run("b"), paragraph.add_run("c")
    on.bold = True
    off.bold = False
    assert (inherit.bold, on.bold, off.bold) == (None, True, False)
    assert "<w:b/>" in to_xml(on.element)
    assert '<w:b w:val="0"/>' in to_xml(off.element)

    again = reopen(doc)
    runs = again.paragraphs[0].runs
    assert [r.bold for r in runs] == [None, True, False]


def test_assigning_none_takes_the_property_back_to_inherited(doc):
    run = doc.add_paragraph().add_run("x")
    run.font.italic = True
    assert run.font.italic is True
    run.font.italic = None
    assert run.font.italic is None
    assert reopen(doc).paragraphs[0].runs[0].font.italic is None


@pytest.mark.parametrize(
    "name",
    [
        "all_caps",
        "complex_script",
        "cs_bold",
        "cs_italic",
        "double_strike",
        "emboss",
        "hidden",
        "imprint",
        "math",
        "no_proof",
        "outline",
        "rtl",
        "shadow",
        "small_caps",
        "snap_to_grid",
        "spec_vanish",
        "strike",
        "web_hidden",
    ],
)
def test_every_boolean_font_property_round_trips_in_all_three_states(doc, name):
    paragraph = doc.add_paragraph()
    runs = [paragraph.add_run("a"), paragraph.add_run("b"), paragraph.add_run("c")]
    setattr(runs[1].font, name, True)
    setattr(runs[2].font, name, False)
    read = [getattr(r.font, name) for r in reopen(doc).paragraphs[0].runs]
    assert read == [None, True, False], name


def test_bold_and_italic_write_their_complex_script_twins(doc):
    from docx4j_py.fragments import to_xml

    run = doc.add_paragraph().add_run("x")
    run.bold = True
    run.italic = True
    xml = to_xml(run.element)
    assert "<w:bCs/>" in xml and "<w:iCs/>" in xml, "CR-003 section 3.12's rule"
    assert run.font.cs_bold is True and run.font.cs_italic is True


def test_subscript_and_superscript_share_one_w_vertalign(doc):
    run = doc.add_paragraph().add_run("x")
    assert run.font.subscript is None and run.font.superscript is None
    run.font.superscript = True
    assert (run.font.superscript, run.font.subscript) == (True, False)
    run.font.subscript = True
    assert (run.font.superscript, run.font.subscript) == (False, True)
    run.font.subscript = False
    assert run.font.subscript is None
    assert reopen(doc).paragraphs[0].runs[0].font.subscript is None


def test_font_name_size_colour_and_highlight(doc):
    run = doc.add_paragraph().add_run("x")
    assert (run.font.name, run.font.size, run.font.color.rgb) == (None, None, None)
    assert run.font.color.type is None
    run.font.name = "Consolas"
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    again = reopen(doc).paragraphs[0].runs[0]
    assert again.font.name == "Consolas"
    assert again.font.size == Pt(14)
    assert again.font.size.pt == 14.0
    assert again.font.color.rgb == RGBColor(0xC0, 0x00, 0x00)
    assert str(again.font.color.rgb) == "C00000"
    assert again.font.color.type.name == "RGB"
    assert again.font.highlight_color is WD_COLOR_INDEX.YELLOW


def test_the_size_is_half_points_in_the_xml_as_word_writes_it(doc):
    from docx4j_py.fragments import to_xml

    run = doc.add_paragraph().add_run("x")
    run.font.size = Pt(14)
    assert '<w:sz w:val="28"/>' in to_xml(run.element)


def test_underline_is_true_false_none_or_a_member(doc):
    paragraph = doc.add_paragraph()
    a, b, c, d = (paragraph.add_run(x) for x in "abcd")
    a.underline = True
    b.underline = False
    c.underline = WD_UNDERLINE.DOUBLE
    assert [r.underline for r in (a, b, c, d)] == [
        True,
        False,
        WD_UNDERLINE.DOUBLE,
        None,
    ]
    read = [r.underline for r in reopen(doc).paragraphs[0].runs]
    assert read == [True, False, WD_UNDERLINE.DOUBLE, None]


def test_a_theme_colour_reads_back_as_one(doc):
    from docx4j_py.docx import MSO_THEME_COLOR

    run = doc.add_paragraph().add_run("x")
    run.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
    again = reopen(doc).paragraphs[0].runs[0]
    assert again.font.color.theme_color is MSO_THEME_COLOR.ACCENT_1
    assert again.font.color.type.name == "THEME"


def test_a_runs_character_style_is_a_style_object(doc):
    run = doc.add_paragraph().add_run("x")
    run.style = "Emphasis"
    assert run.style.name == "Emphasis"
    assert run.style.type is WD_STYLE_TYPE.CHARACTER
    assert reopen(doc).paragraphs[0].runs[0].style.name == "Emphasis"
    run.style = None
    assert run.style.name == "Default Paragraph Font", "the document's default"


def test_run_text_setter_and_add_text_and_add_tab_and_clear(doc):
    run = doc.add_paragraph().add_run()
    run.text = "one"
    run.add_tab()
    run.add_text("two")
    assert run.text == "one\ttwo"
    run.bold = True
    run.clear()
    assert run.text == ""
    assert run.bold is True, "clear keeps the run formatting"


# ---------------------------------------------------------------------------
# ParagraphFormat
# ---------------------------------------------------------------------------


def test_alignment_reads_and_writes_and_none_removes_it(doc):
    paragraph = doc.add_paragraph("x")
    assert paragraph.alignment is None
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    assert paragraph.alignment is WD_ALIGN_PARAGRAPH.CENTER
    assert reopen(doc).paragraphs[0].alignment is WD_ALIGN_PARAGRAPH.CENTER
    paragraph.alignment = None
    assert paragraph.alignment is None


def test_every_alignment_member_this_engine_can_write_round_trips(doc):
    for member in (
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.JUSTIFY,
        WD_ALIGN_PARAGRAPH.DISTRIBUTE,
    ):
        doc.add_paragraph("x").alignment = member
    read = [p.alignment for p in reopen(doc).paragraphs]
    assert read == [
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.JUSTIFY,
        WD_ALIGN_PARAGRAPH.DISTRIBUTE,
    ]


def test_the_indents_and_the_spacing(doc):
    fmt = doc.add_paragraph("x").paragraph_format
    assert (fmt.left_indent, fmt.right_indent, fmt.first_line_indent) == (None, None, None)
    assert (fmt.space_before, fmt.space_after) == (None, None)
    fmt.left_indent = Inches(0.5)
    fmt.right_indent = Cm(1)
    fmt.first_line_indent = Inches(0.25)
    fmt.space_before = Pt(6)
    fmt.space_after = Pt(12)

    again = reopen(doc).paragraphs[0].paragraph_format
    assert again.left_indent == Twips(720)
    assert round(again.right_indent.cm, 2) == 1.0
    assert again.first_line_indent == Twips(360)
    assert again.space_before == Pt(6)
    assert again.space_after == Pt(12)


def test_a_negative_first_line_indent_is_a_hanging_indent(doc):
    from docx4j_py.fragments import to_xml

    fmt = doc.add_paragraph("x").paragraph_format
    fmt.first_line_indent = Inches(-0.25)
    assert 'w:hanging="360"' in to_xml(doc.paragraphs[0].element)
    assert reopen(doc).paragraphs[0].paragraph_format.first_line_indent == Twips(-360)


def test_line_spacing_is_a_float_multiple_or_an_absolute_length(doc):
    multiple = doc.add_paragraph("a").paragraph_format
    exact = doc.add_paragraph("b").paragraph_format
    multiple.line_spacing = 1.5
    exact.line_spacing = Pt(18)

    again = reopen(doc)
    first, second = (p.paragraph_format for p in again.paragraphs)
    assert first.line_spacing == 1.5
    assert first.line_spacing_rule is WD_LINE_SPACING.ONE_POINT_FIVE
    assert second.line_spacing == Pt(18)
    assert second.line_spacing_rule is WD_LINE_SPACING.EXACTLY


def test_the_line_spacing_rule_setter_writes_the_multiple_it_means(doc):
    fmt = doc.add_paragraph("x").paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    assert fmt.line_spacing == 2.0
    assert reopen(doc).paragraphs[0].paragraph_format.line_spacing_rule is (
        WD_LINE_SPACING.DOUBLE
    )


@pytest.mark.parametrize(
    "name", ["keep_together", "keep_with_next", "page_break_before", "widow_control"]
)
def test_the_paragraph_toggles_are_tri_state(doc, name):
    fmt = doc.add_paragraph("x").paragraph_format
    assert getattr(fmt, name) is None
    setattr(fmt, name, True)
    assert getattr(fmt, name) is True
    assert getattr(reopen(doc).paragraphs[0].paragraph_format, name) is True
    setattr(fmt, name, None)
    assert getattr(fmt, name) is None


# ---------------------------------------------------------------------------
# styles
# ---------------------------------------------------------------------------


def test_a_paragraphs_style_is_an_object_with_name_id_type_and_font(doc):
    paragraph = doc.add_heading("Report", 1)
    style = paragraph.style
    assert style.name == "Heading 1"
    assert style.style_id == "Heading1"
    assert style.type is WD_STYLE_TYPE.PARAGRAPH
    assert style.builtin is True
    assert style.font is not None
    assert style.paragraph_format is not None
    assert style.base_style.name == "Normal"
    assert style.next_paragraph_style.name == "Normal"


def test_a_style_can_be_assigned_by_name_by_id_or_as_an_object(doc):
    by_name = doc.add_paragraph("a")
    by_name.style = "Heading 2"
    by_id = doc.add_paragraph("b")
    by_id.style = "Heading2"
    by_object = doc.add_paragraph("c")
    by_object.style = doc.styles["Heading 2"]
    assert [p.style.style_id for p in doc.paragraphs] == ["Heading2"] * 3


def test_assigning_none_removes_the_style(doc):
    paragraph = doc.add_paragraph("x", style="Heading 1")
    paragraph.style = None
    assert paragraph.style.name == "Normal", "the document's default paragraph style"


def test_a_style_the_document_does_not_define_and_word_does_not_know_is_refused(doc):
    from docx4j_py.model.content.errors import StyleError

    with pytest.raises(StyleError):
        doc.add_paragraph("x").style = "Not A Style At All"


def test_list_bullet_is_the_style_and_no_w_numpr(doc):
    """CR-003 section 17.8: python-docx writes a ``w:pStyle`` and no numbering.

    Section 20.9: the *style's* numbering is real, though --- Word paints an
    arbitrary label for a style whose ``w:numId`` the document does not define
    (artefact 11's check showed ``1.`` on a List Bullet paragraph), so the
    definition ``ensure_style`` adds names a ``w:num`` it creates.
    """
    from docx4j_py.fragments import to_xml

    bullet = doc.add_paragraph("a point", style="List Bullet")
    number = doc.add_paragraph("a step", style="List Number")
    for paragraph, style_id in ((bullet, "ListBullet"), (number, "ListNumber")):
        xml = to_xml(paragraph.element)
        assert f'w:pStyle w:val="{style_id}"' in xml
        assert "w:numPr" not in xml
    assert bullet.view.is_list_item is True, "a list through the style, which the emulator resolves"
    assert bullet.view.list_item.list_string == "\u2022"
    assert number.view.list_item.list_string == "1."
    again = reopen(doc)
    assert [p.style.name for p in again.paragraphs] == ["List Bullet", "List Number"]
    assert "List Bullet" in again.styles
    assert [p.view.list_item.list_string for p in again.paragraphs] == ["\u2022", "1."]
    numbering = again.package.numbering_definitions_part.contents
    assert len(numbering.num) == 2, "one w:num per style, as Word's template has"


def test_a_heading_definition_added_here_carries_no_dangling_numbering(doc):
    """Section 20.9: KnownStyles' Heading styles name w:numId 3, which nothing defines."""
    from docx4j_py.fragments import to_xml

    doc.add_heading("Numbered by nothing", level=2)
    styles = doc.package.style_definitions_part.contents
    heading = next(s for s in styles.style if s.style_id == "Heading2")
    assert "numPr" not in to_xml(heading)
    assert doc.package.numbering_definitions_part is None
    assert doc.paragraphs[0].view.is_list_item is False


def test_the_styles_collection_iterates_contains_and_looks_up(doc):
    styles = doc.styles
    assert len(styles) > 5
    assert "Heading 1" in styles
    assert "Not A Style" not in styles
    assert styles["Heading 1"].style_id == "Heading1"
    assert styles["Heading1"].style_id == "Heading1", "an id works too"
    with pytest.raises(KeyError, match="Not A Style"):
        _ = styles["Not A Style"]
    assert styles.default(WD_STYLE_TYPE.PARAGRAPH).name == "Normal"
    names = [style.name for style in styles]
    assert "Normal" in names and len(names) == len(styles)


def test_add_style_and_delete(doc):
    style = doc.styles.add_style("My Own", WD_STYLE_TYPE.PARAGRAPH)
    assert style.style_id == "MyOwn"
    assert style.name == "My Own"
    assert style.builtin is False
    style.font.bold = True
    style.paragraph_format.space_after = Pt(18)
    doc.add_paragraph("x").style = "My Own"

    again = reopen(doc)
    assert again.styles["My Own"].font.bold is True
    assert again.styles["My Own"].paragraph_format.space_after == Pt(18)
    assert again.paragraphs[0].style.name == "My Own"
    with pytest.raises(ValueError, match="already contains"):
        again.styles.add_style("My Own", WD_STYLE_TYPE.PARAGRAPH)

    again.styles["My Own"].delete()
    assert "My Own" not in again.styles


def test_a_styles_hidden_priority_and_quick_style_round_trip(doc):
    style = doc.styles.add_style("Quiet", WD_STYLE_TYPE.PARAGRAPH)
    style.hidden = True
    style.priority = 42
    style.quick_style = True
    style.unhide_when_used = True
    again = reopen(doc).styles["Quiet"]
    assert (again.hidden, again.priority, again.quick_style, again.unhide_when_used) == (
        True,
        42,
        True,
        True,
    )


# ---------------------------------------------------------------------------
# tables
# ---------------------------------------------------------------------------


def test_add_table_sizes_the_grid_from_the_section(doc):
    table = doc.add_table(2, 3)
    assert len(table.rows) == 2
    assert len(table.columns) == 3
    widths = [column.width for column in table.columns]
    assert all(width is not None for width in widths)
    section = doc.sections[0]
    writable = section.page_width - section.left_margin - section.right_margin
    assert abs(sum(widths) - writable) <= 3, "the grid sums to the text width"


def test_cells_read_and_write_by_grid_position(doc):
    table = doc.add_table(2, 2)
    table.cell(0, 0).text = "a"
    table.cell(0, 1).text = "b"
    table.cell(1, 0).text = "c"
    table.cell(1, 1).text = "d"
    assert [c.text for c in table.rows[0].cells] == ["a", "b"]
    assert [c.text for c in table.columns[1].cells] == ["b", "d"]
    assert [c.text for c in table.row_cells(1)] == ["c", "d"]
    assert [c.text for c in table.column_cells(0)] == ["a", "c"]
    with pytest.raises(IndexError):
        table.cell(0, 2)

    again = reopen(doc).tables[0]
    assert [[c.text for c in row.cells] for row in again.rows] == [["a", "b"], ["c", "d"]]


def test_a_cells_text_setter_replaces_everything_with_one_paragraph(doc):
    cell = doc.add_table(1, 1).cell(0, 0)
    cell.add_paragraph("one")
    cell.add_paragraph("two")
    assert cell.text == "\none\ntwo"
    cell.text = "just this"
    assert cell.text == "just this"
    assert len(cell.paragraphs) == 1


def test_add_row_and_add_column(doc):
    table = doc.add_table(1, 2)
    table.cell(0, 0).text = "a"
    row = table.add_row()
    assert len(table.rows) == 2
    row.cells[0].text = "b"
    column = table.add_column(Inches(1))
    assert len(table.columns) == 3
    assert column.width == Inches(1)
    table.cell(0, 2).text = "c"

    again = reopen(doc).tables[0]
    assert [[c.text for c in r.cells] for r in again.rows] == [["a", "", "c"], ["b", "", ""]]
    assert again.columns[2].width == Inches(1)


def test_a_grid_span_cell_appears_at_every_position_it_covers(doc):
    """python-docx's layout-grid rule: a merged cell repeats."""
    doc.part.body.insert_xml(
        "<w:tbl><w:tblPr/>"
        '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>'
        '<w:gridCol w:w="2000"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'
        "<w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc>"
        '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>'
        "<w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc></w:tr>"
        "<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>"
    )
    table = reopen(doc).tables[0]
    assert [[table.cell(r, c).text for c in range(3)] for r in range(2)] == [
        ["wide", "wide", "tall"],
        ["a", "b", "tall"],
    ]
    assert table.cell(0, 0) == table.cell(0, 1), "one cell, two grid positions"
    assert table.cell(0, 2) == table.cell(1, 2), "the vertical span answers upwards"
    assert table.cell(0, 0).grid_span == 2
    assert table.cell(1, 0).grid_span == 1
    assert [c.text for c in table.rows[1].cells] == ["a", "b", "tall"]


def test_a_rows_grid_cols_before_and_after(doc):
    doc.part.body.insert_xml(
        "<w:tbl><w:tblPr/>"
        '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
        '<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>'
        "<w:tc><w:p><w:r><w:t>late</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    )
    row = reopen(doc).tables[0].rows[0]
    assert row.grid_cols_before == 1
    assert row.grid_cols_after == 0
    assert [c.text for c in row.cells] == ["late"]


def test_a_tables_style_alignment_autofit_and_direction(doc):
    from docx4j_py.docx import WD_TABLE_DIRECTION

    table = doc.add_table(1, 1)
    assert table.style.name == "Normal Table"
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.table_direction = WD_TABLE_DIRECTION.RTL

    again = reopen(doc).tables[0]
    assert again.style.name == "Table Grid"
    assert again.alignment is WD_TABLE_ALIGNMENT.CENTER
    assert again.autofit is False
    assert again.table_direction is WD_TABLE_DIRECTION.RTL


def test_a_rows_height_and_rule_and_a_cells_width_and_alignment(doc):
    table = doc.add_table(1, 1)
    table.rows[0].height = Inches(0.5)
    table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    table.cell(0, 0).width = Inches(2)
    table.cell(0, 0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    again = reopen(doc).tables[0]
    assert again.rows[0].height == Inches(0.5)
    assert again.rows[0].height_rule is WD_ROW_HEIGHT_RULE.EXACTLY
    assert again.cell(0, 0).width == Inches(2)
    assert again.cell(0, 0).vertical_alignment is WD_ALIGN_VERTICAL.CENTER


def test_a_nested_table_in_a_cell_and_the_paragraph_after_it(doc):
    cell = doc.add_table(1, 1).cell(0, 0)
    inner = cell.add_table(2, 2)
    inner.cell(0, 0).text = "in"
    assert len(cell.tables) == 1
    assert cell.paragraphs[-1].text == "", "Word wants a paragraph last in a cell"

    again = reopen(doc)
    assert len(again.tables) == 1, "only top-level tables are in document.tables"
    outer_cell = again.tables[0].cell(0, 0)
    assert len(outer_cell.tables) == 1
    assert outer_cell.tables[0].cell(0, 0).text == "in"


def test_add_table_with_no_rows_is_a_table_to_add_rows_to(doc):
    table = doc.add_table(0, 3)
    assert len(table.rows) == 0
    assert len(table.columns) == 3
    table.add_row().cells[0].text = "first"
    assert [c.text for c in reopen(doc).tables[0].rows[0].cells] == ["first", "", ""]
    with pytest.raises(ValueError, match="one column"):
        doc.add_table(1, 0)


# ---------------------------------------------------------------------------
# sections, headers and footers
# ---------------------------------------------------------------------------


def test_the_page_setup_reads_and_writes(doc):
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.header_distance = Inches(0.5)
    section.footer_distance = Inches(0.5)
    section.gutter = Inches(0.25)

    again = reopen(doc).sections[0]
    assert again.page_width.inches == 8.5
    assert again.page_height.inches == 11.0
    assert again.left_margin == Inches(1)
    assert again.header_distance == Inches(0.5)
    assert again.footer_distance == Inches(0.5)
    assert again.gutter == Inches(0.25)


def test_add_section_clones_the_page_setup_and_sets_the_start_type(doc):
    doc.add_paragraph("first section")
    second = doc.add_section(WD_SECTION.NEW_PAGE)
    assert len(doc.sections) == 2
    assert second.start_type is WD_SECTION.NEW_PAGE
    second.orientation = WD_ORIENT.LANDSCAPE
    second.page_width, second.page_height = Inches(11.69), Inches(8.27)

    again = reopen(doc)
    assert len(again.sections) == 2
    assert again.sections[0].orientation is WD_ORIENT.PORTRAIT
    assert again.sections[1].orientation is WD_ORIENT.LANDSCAPE
    assert round(again.sections[1].page_width.inches, 2) == 11.69


def test_every_start_type_round_trips(doc):
    for member in (
        WD_SECTION.CONTINUOUS,
        WD_SECTION.NEW_COLUMN,
        WD_SECTION.EVEN_PAGE,
        WD_SECTION.ODD_PAGE,
    ):
        doc.add_paragraph("x")
        doc.add_section(member)
    # python-docx's rule: add_section sets the type on the new **last** section,
    # and the clone left behind carries what the previous sentinel said
    read = [section.start_type for section in reopen(doc).sections]
    assert read == [
        WD_SECTION.NEW_PAGE,
        WD_SECTION.CONTINUOUS,
        WD_SECTION.NEW_COLUMN,
        WD_SECTION.EVEN_PAGE,
        WD_SECTION.ODD_PAGE,
    ]


def test_a_sections_inner_content_is_the_blocks_it_covers(doc):
    doc.add_paragraph("one")
    doc.add_section()
    doc.add_paragraph("two")
    first, second = doc.sections
    assert [p.text for p in first.iter_inner_content()] == ["one", ""]
    assert [p.text for p in second.iter_inner_content()] == ["two"]


def test_a_header_is_created_by_unlinking_it_and_carries_its_paragraph(doc):
    header = doc.sections[0].header
    assert header.is_linked_to_previous is True
    assert doc.package.header_parts() == []

    header.is_linked_to_previous = False
    assert header.is_linked_to_previous is False
    assert len(doc.package.header_parts()) == 1
    header.paragraphs[0].text = "docx4j-python"
    header.add_paragraph("and a second line")

    footer = doc.sections[0].footer
    footer.is_linked_to_previous = False
    footer.paragraphs[0].text = "page"

    data = doc.save()
    names = set(part_bytes(data, relationships=True))
    assert "word/header1.xml" in names
    assert "word/footer1.xml" in names
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        types = archive.read("[Content_Types].xml").decode()
        assert "wordprocessingml.header+xml" in types
        assert "wordprocessingml.footer+xml" in types

    again = Document(io.BytesIO(data)).sections[0]
    assert again.header.is_linked_to_previous is False
    assert [p.text for p in again.header.paragraphs] == [
        "docx4j-python",
        "and a second line",
    ]
    assert again.footer.paragraphs[0].text == "page"


def test_relinking_a_header_removes_the_reference(doc):
    header = doc.sections[0].header
    header.is_linked_to_previous = False
    header.paragraphs[0].text = "gone"
    header.is_linked_to_previous = True
    assert header.is_linked_to_previous is True
    from docx4j_py.fragments import to_xml

    assert "headerReference" not in to_xml(doc.sections[0].element)
    again = reopen(doc)
    assert again.sections[0].header.is_linked_to_previous is True


def test_a_second_section_inherits_the_first_sections_header(doc):
    first = doc.sections[0]
    first.header.is_linked_to_previous = False
    first.header.paragraphs[0].text = "shared"
    doc.add_paragraph("x")
    doc.add_section()

    again = reopen(doc)
    assert len(again.sections) == 2
    assert again.sections[1].header.is_linked_to_previous is True
    assert again.sections[1].header.paragraphs[0].text == "shared", "inherited"


def test_the_three_header_kinds_are_three_parts(doc):
    section = doc.sections[0]
    section.different_first_page_header_footer = True
    section.header.is_linked_to_previous = False
    section.first_page_header.is_linked_to_previous = False
    section.even_page_header.is_linked_to_previous = False
    section.header.paragraphs[0].text = "default"
    section.first_page_header.paragraphs[0].text = "first"
    section.even_page_header.paragraphs[0].text = "even"

    again = reopen(doc).sections[0]
    assert again.different_first_page_header_footer is True
    assert again.header.paragraphs[0].text == "default"
    assert again.first_page_header.paragraphs[0].text == "first"
    assert again.even_page_header.paragraphs[0].text == "even"


def test_a_header_can_hold_a_table(doc):
    header = doc.sections[0].header
    header.is_linked_to_previous = False
    header.add_table(1, 2, Inches(4)).cell(0, 0).text = "left"
    again = reopen(doc).sections[0].header
    assert len(again.tables) == 1
    assert again.tables[0].cell(0, 0).text == "left"


def test_a_header_read_from_a_sample_says_what_it_is(doc):
    document = Document(SAMPLES / "Headers.docx")
    section = document.sections[0]
    assert section.header.is_linked_to_previous is False
    assert any(p.text for p in section.header.paragraphs)


# ---------------------------------------------------------------------------
# pictures
# ---------------------------------------------------------------------------


def test_add_picture_returns_an_inline_shape_and_writes_the_part(doc, png):
    shape = doc.add_picture(png, width=Inches(2))
    assert shape.width == Inches(2)
    assert shape.height > 0
    assert shape.type.name == "PICTURE"

    data = doc.save()
    names = set(part_bytes(data))
    assert "word/media/image1.png" in names

    again = Document(io.BytesIO(data))
    assert len(again.inline_shapes) == 1
    assert again.inline_shapes[0].width == Inches(2)
    assert again.inline_shapes[0].type.name == "PICTURE"


def test_the_shapes_size_can_be_set_afterwards(doc, png):
    shape = doc.add_picture(png, width=Inches(2))
    shape.width = Inches(3)
    shape.height = Inches(1)
    again = reopen(doc).inline_shapes[0]
    assert again.width == Inches(3)
    assert again.height == Inches(1), "python-docx sets only the axis it is given"


def test_a_picture_in_a_run_and_in_a_paragraph(doc, png):
    paragraph = doc.add_paragraph("before ")
    shape = paragraph.add_run().add_picture(png, height=Inches(1))
    assert shape.height == Inches(1)
    assert paragraph.text == "before ", "a drawing contributes no text"
    again = reopen(doc)
    assert len(again.inline_shapes) == 1
    assert again.paragraphs[0].text == "before "


def test_add_picture_takes_a_path_and_a_stream(doc, tmp_path, png):
    path = tmp_path / "pangolin.png"
    path.write_bytes(png)
    doc.add_picture(str(path), width=Inches(1))
    doc.add_picture(path, width=Inches(1))
    doc.add_picture(io.BytesIO(png), width=Inches(1))
    again = reopen(doc)
    assert len(again.inline_shapes) == 3
    assert {shape.width for shape in again.inline_shapes} == {Inches(1)}


def test_a_runs_inner_content_names_its_text_breaks_and_drawings(doc, png):
    run = doc.add_paragraph().add_run("a")
    run.add_tab()
    run.add_text("b")
    run.add_picture(png, width=Inches(1))
    kinds = [
        item if isinstance(item, str) else type(item).__name__
        for item in run.iter_inner_content()
    ]
    assert kinds == ["a\tb", "Drawing"]
    drawing = next(item for item in run.iter_inner_content() if not isinstance(item, str))
    assert drawing.has_picture is True


# ---------------------------------------------------------------------------
# core properties and settings
# ---------------------------------------------------------------------------


def test_the_core_properties_read_and_write_and_round_trip(doc):
    properties = doc.core_properties
    properties.author = "Ada Lovelace"
    properties.title = "Quarterly report"
    properties.subject = "Numbers"
    properties.keywords = "quarter, revenue"
    properties.comments = "A description"
    properties.category = "Reports"
    properties.content_status = "Draft"
    properties.identifier = "urn:report:1"
    properties.language = "en-GB"
    properties.last_modified_by = "Bob Bones"
    properties.version = "1.0"
    properties.revision = 3
    properties.created = dt.datetime(2026, 9, 18, 10, 0, tzinfo=dt.UTC)
    properties.modified = dt.datetime(2026, 9, 18, 11, 0, tzinfo=dt.UTC)
    properties.last_printed = dt.datetime(2026, 9, 1, 9, 0, tzinfo=dt.UTC)

    again = reopen(doc).core_properties
    assert again.author == "Ada Lovelace"
    assert again.title == "Quarterly report"
    assert again.subject == "Numbers"
    assert again.keywords == "quarter, revenue"
    assert again.comments == "A description"
    assert again.category == "Reports"
    assert again.content_status == "Draft"
    assert again.identifier == "urn:report:1"
    assert again.language == "en-GB"
    assert again.last_modified_by == "Bob Bones"
    assert again.version == "1.0"
    assert again.revision == 3
    assert again.created == dt.datetime(2026, 9, 18, 10, 0, tzinfo=dt.UTC)
    assert again.modified == dt.datetime(2026, 9, 18, 11, 0, tzinfo=dt.UTC)
    assert again.last_printed == dt.datetime(2026, 9, 1, 9, 0, tzinfo=dt.UTC)


def test_an_unset_core_property_is_the_empty_string_as_python_docx_reports_it(doc):
    assert doc.core_properties.title == ""
    assert doc.core_properties.author == ""
    assert doc.core_properties.revision == 0
    assert doc.core_properties.last_printed is None


def test_a_loaded_documents_properties_are_read_as_they_are():
    properties = Document(SAMPLES / "2010-sample1.docx").core_properties
    assert properties.author == "jharrop"
    assert properties.last_modified_by == "jharrop"
    assert properties.revision == 2
    assert properties.created.year == 2011


@pytest.mark.filterwarnings("ignore::docx4j_xsdata.exceptions.ConverterWarning")
def test_the_core_dates_are_aware_in_utc_in_every_form_as_python_docx_reads_them():
    """python-docx 1.2.0's ``_parse_W3CDTF_to_datetime``: UTC whatever the form, as Word reads a core property's.

    docx4j's date-forms probe (CR-067 section 12.2; CR-025 section 13.10 item 2):
    ``created`` a date alone, ``modified`` a time with no offset, ``lastPrinted``
    with a ``Z``. They were naive for the first two.
    """
    properties = Document(SAMPLES.parent / "tests" / "fixtures" / "date-forms-cr067" / "date-forms-r1.docx").core_properties
    assert properties.created == dt.datetime(2024, 2, 3, tzinfo=dt.UTC)
    assert properties.modified == dt.datetime(2025, 6, 7, 8, 9, 10, tzinfo=dt.UTC)
    assert properties.last_printed == dt.datetime(2025, 1, 2, 3, 4, 5, tzinfo=dt.UTC)
    assert all(d.tzinfo is dt.UTC for d in (properties.created, properties.modified, properties.last_printed))


def test_the_revision_setter_refuses_what_word_refuses(doc):
    with pytest.raises(ValueError, match="positive int"):
        doc.core_properties.revision = 0


def test_the_one_setting_python_docx_offers(doc):
    assert doc.settings.odd_and_even_pages_header_footer is False
    doc.settings.odd_and_even_pages_header_footer = True
    assert reopen(doc).settings.odd_and_even_pages_header_footer is True


# ---------------------------------------------------------------------------
# comments
# ---------------------------------------------------------------------------


def test_add_comment_anchors_on_the_runs_and_creates_the_parts(doc):
    paragraph = doc.add_paragraph("The ")
    anchor = paragraph.add_run("quick brown fox")
    paragraph.add_run(" jumps")
    comment = doc.add_comment(anchor, "Is this right?", author="Ada Lovelace", initials="AL")
    assert comment.text == "Is this right?"
    assert comment.author == "Ada Lovelace"
    assert comment.initials == "AL"
    assert comment.comment_id == 0
    assert comment.timestamp is not None

    data = doc.save()
    names = set(part_bytes(data))
    assert "word/comments.xml" in names
    assert "word/commentsExtended.xml" in names

    again = Document(io.BytesIO(data))
    assert len(again.comments) == 1
    first = next(iter(again.comments))
    assert first.text == "Is this right?"
    assert first.author == "Ada Lovelace"
    assert again.comments.get(0).text == "Is this right?"
    assert again.comments.get(99) is None
    # the anchor is exactly the run that was given
    ranges = first.view.get_ranges()  # Comment.get_range() is one Range since CR-013 Phase B (D2); get_ranges() the old list
    assert [str(span) for span in ranges] == ["quick brown fox"]


def test_a_comment_can_be_given_a_whole_paragraphs_runs(doc):
    paragraph = doc.add_paragraph("one ")
    paragraph.add_run("two ")
    paragraph.add_run("three")
    comment = doc.add_comment(paragraph.runs, "all of it")
    assert [str(span) for span in comment.view.get_ranges()] == ["one two three"]  # Comment.get_range() is one Range since CR-013 Phase B (D2); get_ranges() the old list


def test_a_comment_takes_more_paragraphs_and_keeps_the_comment_style(doc):
    run = doc.add_paragraph("x").add_run("y")
    comment = doc.add_comment(run, "first line")
    comment.add_paragraph("second line")
    assert comment.text == "first line\nsecond line"
    assert comment.paragraphs[1].style.name == "Comment Text"
    again = next(iter(reopen(doc).comments))
    assert again.text == "first line\nsecond line"


def test_a_comment_needs_runs_and_refuses_to_span_paragraphs(doc):
    from docx4j_py.model.content.errors import ContentError

    first = doc.add_paragraph("a").add_run("b")
    second = doc.add_paragraph("c").add_run("d")
    with pytest.raises(ContentError, match="at least one run"):
        doc.add_comment([], "nothing")
    with pytest.raises(ContentError, match="same paragraph"):
        doc.add_comment([first, second], "across")


def test_the_authors_and_initials_survive_and_the_package_author_is_restored(doc):
    from docx4j_py.model.content.comments import Author

    doc.package.author = Author("Claude", initials="C")
    run = doc.add_paragraph("x").add_run("y")
    doc.add_comment(run, "by Ada", author="Ada Lovelace", initials="AL")
    assert doc.package.author.name == "Claude", "the package identity is put back"
    second = doc.add_paragraph("p").add_run("q")
    doc.add_comment(second, "by the package author")
    authors = [comment.author for comment in reopen(doc).comments]
    assert authors == ["Ada Lovelace", "Claude"]


# ---------------------------------------------------------------------------
# what is underneath
# ---------------------------------------------------------------------------


def test_the_typed_tree_and_the_content_api_stay_reachable(doc):
    from docx4j_py.model.content import Paragraph as ContentParagraph
    from docx4j_py.wml import P, R

    paragraph = doc.add_paragraph("x")
    assert isinstance(paragraph.element, P)
    assert isinstance(paragraph._element, P)
    assert isinstance(paragraph.runs[0].element, R)
    assert isinstance(paragraph.view, ContentParagraph)
    assert paragraph.view.address
    assert doc.part.body.text == "x"
    assert doc.package is doc.part.package
    assert paragraph.part is doc.part


def test_every_facade_mutation_leaves_a_change_report(doc):
    doc.package.changes.clear()
    doc.add_heading("Report", 1)
    doc.add_paragraph("body").add_run("more").bold = True
    doc.add_table(1, 1).cell(0, 0).text = "cell"
    operations = [change.operation for change in doc.package.changes]
    assert operations, "a facade call records a ChangeReport like any other"
    assert all(change.parts_touched for change in doc.package.changes)
    assert any(operation.startswith("docx.") for operation in operations)
    assert "/word/document.xml" in doc.package.last_change.parts_touched


def test_a_facade_write_can_be_tried_and_thrown_away(doc, png):
    """``pkg.dry_run()`` covers a facade call, because the facade is over the API.

    The facade holds a package, so a trial one is what a trial ``Document``
    needs; ``on_trial`` below is the two lines an MCP server would write.
    """
    doc.add_paragraph("keep")
    before = doc.save()
    with doc.package.dry_run() as trial:
        trial_document = on_trial(trial)
        trial_document.add_paragraph("thrown away")
        trial_document.add_picture(png, width=Inches(1))
        assert [p.text for p in trial_document.paragraphs] == [
            "keep",
            "thrown away",
            "",
        ]
    assert [p.text for p in doc.paragraphs] == ["keep"]
    assert part_bytes(doc.save()) == part_bytes(before), "the picture part went too"


def on_trial(trial: object) -> Document:
    """A facade ``Document`` over a trial package, for ``dry_run``."""
    from docx4j_py.docx._base import Story
    from docx4j_py.model.content.body import body_of

    document = Document.__new__(Document)
    document._package = trial
    document._body_story = Story(body_of(trial.get_main_document_part()), document)
    return document


def test_two_identical_scripts_give_identical_bytes(png):
    def build() -> bytes:
        document = Document()
        document.package.id_seed = SEED
        document.core_properties.created = dt.datetime(
            2026, 9, 18, 10, 0, tzinfo=dt.UTC
        )
        document.core_properties.modified = document.core_properties.created
        document.add_heading("Report", 1)
        document.add_paragraph("Body text").add_run(" bold").bold = True
        document.add_table(2, 2, style="Table Grid").cell(0, 0).text = "a"
        document.add_picture(png, width=Inches(2))
        return document.save()

    assert part_bytes(build()) == part_bytes(build())


# -- tab stops (added 2026-09-18) ------------------------------------------


def test_tab_stops_are_added_in_position_order_with_python_docxs_defaults(doc):
    from docx4j_py.docx import WD_TAB_ALIGNMENT, WD_TAB_LEADER
    from docx4j_py.fragments import to_xml

    paragraph = doc.add_paragraph("Name\tPage")
    stops = paragraph.paragraph_format.tab_stops
    assert len(stops) == 0
    right = stops.add_tab_stop(Inches(3), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    left = stops.add_tab_stop(Inches(1))
    assert [s.position.inches for s in stops] == [1.0, 3.0], "sorted by position"
    assert left.alignment is WD_TAB_ALIGNMENT.LEFT and left.leader is WD_TAB_LEADER.SPACES
    assert right.alignment is WD_TAB_ALIGNMENT.RIGHT and right.leader is WD_TAB_LEADER.DOTS
    xml = to_xml(paragraph._p.p_pr)
    assert '<w:tab w:val="left" w:pos="1440"/>' in xml, "SPACES writes no w:leader"
    assert '<w:tab w:val="right" w:leader="dot" w:pos="4320"/>' in xml
    assert tuple(doc.package.last_change.parts_touched) == ("/word/document.xml",)


def test_a_tab_stop_moves_and_keeps_the_order_and_deleting_the_last_removes_w_tabs(doc):
    from docx4j_py.docx import WD_TAB_ALIGNMENT
    from docx4j_py.fragments import to_xml

    paragraph = doc.add_paragraph("x")
    stops = paragraph.paragraph_format.tab_stops
    stops.add_tab_stop(Inches(1))
    second = stops.add_tab_stop(Inches(2), WD_TAB_ALIGNMENT.CENTER)
    second.position = Inches(0.5)
    assert [s.position.inches for s in stops] == [0.5, 1.0]
    assert stops[0].alignment is WD_TAB_ALIGNMENT.CENTER
    assert stops[-1].position == Twips(1440)
    del stops[0]
    assert len(stops) == 1
    del stops[0]
    assert "w:tabs" not in to_xml(paragraph._p)
    with pytest.raises(IndexError):
        stops[0]
    stops.add_tab_stop(Inches(1))
    stops.clear_all()
    assert "w:tabs" not in to_xml(paragraph._p)
    stops.clear_all()  # nothing there: nothing recorded, nothing raised


def test_tab_stops_survive_a_reload_and_a_style_has_them_too(doc):
    from docx4j_py.docx import WD_STYLE_TYPE, WD_TAB_ALIGNMENT, WD_TAB_LEADER

    paragraph = doc.add_paragraph("a")
    paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(-0.5), WD_TAB_ALIGNMENT.DECIMAL, WD_TAB_LEADER.DASHES)
    style = doc.styles.add_style("Ledger", WD_STYLE_TYPE.PARAGRAPH)
    style.paragraph_format.tab_stops.add_tab_stop(Inches(4), WD_TAB_ALIGNMENT.RIGHT)
    again = reopen(doc)
    stop = again.paragraphs[0].paragraph_format.tab_stops[0]
    assert (stop.position.inches, stop.alignment, stop.leader) == (-0.5, WD_TAB_ALIGNMENT.DECIMAL, WD_TAB_LEADER.DASHES)
    ledger = again.styles["Ledger"].paragraph_format.tab_stops
    assert [(s.position.inches, s.alignment) for s in ledger] == [(4.0, WD_TAB_ALIGNMENT.RIGHT)]
