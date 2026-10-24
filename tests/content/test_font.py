"""``Font``: the Office JS run properties, and the formatting it reads.

CR-003 section 6, closed by CR-002 Phase B (section 15.1): reads report the
**effective** formatting, as Office JS does --- document defaults, paragraph
style, character style, direct --- and ``get_font(direct=True)`` reads the
runs' own ``w:rPr``, which is what these reads were before the resolver landed.
Since CR-016 Phase B every run in scope is read, ``None``, ``""`` or
``"Mixed"`` where they differ, an empty paragraph reads its mark and an
automatic colour ``#000000`` (core-ts ``content.test.mjs`` at ``222cc20``, its
two font tests ported below with their strings kept).
"""

from __future__ import annotations

from conftest import reloaded, sample

from docx4j_py.wml import to_xml


def test_every_property_writes_and_reads_back(new_package):
    paragraph = new_package.body.insert_paragraph("formatted")
    font = paragraph.font

    font.bold = True
    font.italic = True
    font.strike_through = True
    font.underline = "Double"
    font.name = "Arial"
    font.size = 14
    font.color = "#FF0000"
    font.highlight_color = "yellow"

    assert font.bold is True
    assert font.italic is True
    assert font.strike_through is True
    assert font.underline == "Double"
    assert font.name == "Arial"
    assert font.size == 14
    assert font.color == "#FF0000"
    assert font.highlight_color == "#FFFF00"

    r_pr = paragraph.runs[0].r_pr
    assert r_pr.b is not None and r_pr.b_cs is not None
    assert r_pr.i is not None and r_pr.i_cs is not None
    assert r_pr.sz.val == 28 and r_pr.sz_cs.val == 28, "the szCs twin, CR-003 section 3.12"
    assert r_pr.color.val == "FF0000"
    assert r_pr.highlight.val == "yellow"
    assert r_pr.r_fonts.ascii == "Arial" and r_pr.r_fonts.h_ansi == "Arial"

    back = reloaded(new_package).body.paragraphs[0].font
    assert back.to_dict() == font.to_dict()


def test_subscript_and_superscript_share_w_vert_align(new_package):
    font = new_package.body.insert_paragraph("x").font

    font.superscript = True
    assert (font.superscript, font.subscript) == (True, False)
    font.subscript = True
    assert (font.superscript, font.subscript) == (False, True)
    font.subscript = False
    assert (font.superscript, font.subscript) == (False, False)


def test_the_empty_values_remove_the_element(new_package):
    paragraph = new_package.body.insert_paragraph("x")
    font = paragraph.font
    font.bold = True
    font.size = 12
    font.name = "Arial"
    font.color = "00FF00"
    font.highlight_color = "cyan"
    font.underline = "Single"

    assert font.color == "#00FF00", "the # is optional on the way in"

    font.bold = False
    font.size = 0
    font.name = ""
    font.color = ""
    font.highlight_color = None
    font.underline = "None"

    r_pr = paragraph.runs[0].r_pr
    assert r_pr.b is None
    assert r_pr.sz is None and r_pr.sz_cs is None
    assert r_pr.r_fonts is None
    assert r_pr.color is None
    assert r_pr.highlight is None
    assert r_pr.u is None
    assert paragraph.get_font(direct=True).to_dict() == {
        "bold": False,
        "italic": False,
        "underline": "None",
        "strike_through": False,
        "double_strike_through": False,
        "subscript": False,
        "superscript": False,
        "name": "",
        "size": 0,
        "color": "",
        "highlight_color": None,
        "style": "",
        # w:color's theme attributes as written (CR-009 Phase B)
        "theme_color": None,
        "theme_tint": None,
        "theme_shade": None,
    }
    # the effective read still reports the document defaults underneath
    assert font.size == 11.0


def test_a_write_reaches_every_run_a_read_reads_every_run(new_package):
    paragraph = new_package.body.insert_paragraph("plain bold plain")
    paragraph.search("bold")[0].font.bold = True
    assert len(paragraph.runs) == 3

    # CR-016 Phase B (core-ts CR-002 section 59): every run is read, None where they differ;
    # until then the first run's, False
    assert paragraph.font.bold is None, "the runs differ"
    paragraph.font.name = "Arial"
    assert all(run.r_pr.r_fonts.ascii == "Arial" for run in paragraph.runs)


def test_repr_says_what_the_font_is(new_package):
    paragraph = new_package.body.insert_paragraph("x")
    assert repr(paragraph.get_font(direct=True)) == "<Font paragraph plain>"
    # effective, the document defaults give it a size and the theme a face, and an automatic
    # colour reads #000000, as Office JS reads it (CR-016 Phase B)
    assert repr(paragraph.font) == "<Font paragraph Aptos 11.0pt #000000>"
    paragraph.font.bold = True
    paragraph.font.size = 11
    assert repr(paragraph.font) == "<Font paragraph bold Aptos 11.0pt #000000>"


def test_font_reads_direct_formatting_of_a_run_in_a_real_document():
    package = sample("w14_texteffects.docx")
    paragraph = package.body.paragraphs[0]
    assert paragraph.text == "Testing"
    # whatever the run says directly is what is reported
    assert paragraph.get_font(direct=True).to_dict()["size"] == (
        paragraph.get_font(direct=True).size
    )


def test_font_reports_the_effective_formatting_of_a_styled_paragraph(new_package):
    """Office JS reports effective formatting, and so does this (CR-002 §15.1).

    ``Heading 1`` in docx4j's default styles is bold and 14 pt; a paragraph in
    that style with no direct run formatting therefore reads ``bold=True`` and
    ``size=14``, where the direct read reports nothing at all. This test was
    ``xfail`` until the resolver landed, as CR-003 section 6 said it would be.
    """
    paragraph = new_package.body.insert_paragraph("Heading", style="Heading 1")
    assert paragraph.runs[0].r_pr is None, "nothing is set directly"
    assert paragraph.font.size == 14
    assert paragraph.font.bold is True
    direct = paragraph.get_font(direct=True)
    assert direct.size == 0 and direct.bold is False


# ---------------------------------------------------------------------------
# CR-016 Phase B: every run read (core-ts content.test.mjs at 222cc20, its CR-002 section 59)
# ---------------------------------------------------------------------------


def test_font_reads_direct_formatting_writes_to_runs_splits_runs_for_a_range(new_package):
    """core-ts's 'font: reads direct formatting, writes to runs, splits runs for a range'."""
    p = new_package.body.insert_paragraph("plain bold plain")
    span = p.search("bold")[0]
    span.font.bold = True
    span.font.italic = True
    span.font.size = 14
    span.font.color = "#FF0000"
    span.font.highlight_color = "yellow"
    span.font.underline = "Single"
    assert len(p.runs) == 3, "run split at both ends"
    assert p.text == "plain bold plain"
    bold = p.runs[1]
    assert bold.r_pr.sz.val == 28
    assert span.font.bold is True
    assert span.font.size == 14
    assert span.font.color == "#FF0000"
    assert span.font.highlight_color == "#FFFF00"
    assert span.font.underline == "Single"
    assert p.runs[0].r_pr is None, "other runs untouched"
    # the paragraph's runs differ: None, '' or Mixed, as Office JS reads them (CR-002 section 59, checks 52 and 53)
    font = p.font
    assert [font.bold, font.italic, font.size, font.color, font.highlight_color, font.underline] == [None, None, None, "", "", "Mixed"]
    new_package.get_run_font_selector()
    assert p.font.name == "Aptos", "where they agree, their value"
    assert p.get_font(direct=True).bold is None, "the direct view reads the same way"
    p.font.name = "Arial"
    assert all(run.r_pr.r_fonts.ascii == "Arial" for run in p.runs)
    span.font.bold = False
    assert bold.r_pr.b is None
    assert "<w:i/>" in to_xml(p.element)
    # saved text survives the split
    assert reloaded(new_package).body.paragraphs[0].text == "plain bold plain"


def test_font_reads_as_office_js_an_automatic_colour_000000_an_empty_paragraph_its_mark_the_text_of_words_default_view(new_package):
    """core-ts's 'font reads as Office JS: an automatic colour #000000, an empty paragraph its mark, the text of Word's default view (CR-002 section 59)'."""
    body = new_package.body
    body.insert_xml(
        "<w:p><w:r><w:t>plain</w:t></w:r></w:p>"
        '<w:p><w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>auto</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:b w:val="0"/></w:rPr></w:pPr></w:p>'
        '<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r><w:del w:id="1" w:author="a"><w:r><w:rPr><w:b/></w:rPr>'
        "<w:delText>gone</w:delText></w:r></w:del></w:p>",
        location="End",
    )
    plain, auto, heading, mixed = body.paragraphs[-4:]
    # no colour anywhere, and w:val="auto": #000000, as check 53 reads every run with no colour; the direct view says what is written
    assert [plain.font.color, auto.font.color] == ["#000000", "#000000"]
    assert [plain.get_font(direct=True).color, auto.get_font(direct=True).color] == ["", "auto"]
    # a paragraph with no text reads its mark's w:rPr, over its style's bold (check 52, case 60), its whole range too
    assert heading.style == "Heading 1"
    assert heading.font.bold is False
    assert heading.get_range("Whole").font.bold is False
    assert heading.get_range("Start").font.bold is True, "a collapsed range reads the style, as before"
    # the deleted bold run isn't in Paragraph.text, so it isn't read: one run, not bold
    assert mixed.text == "kept "
    assert mixed.font.bold is False


def test_the_mark_read_leaves_out_the_marks_own_revisions_and_a_placeholders_runs_are_read_where_there_is_no_text(new_package):
    body = new_package.body
    body.insert_xml(
        '<w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="a"/><w:i/><w:rPrChange w:id="4" w:author="a"><w:rPr><w:b/></w:rPr>'
        "</w:rPrChange></w:rPr></w:pPr></w:p>"
        '<w:sdt><w:sdtPr><w:id w:val="9"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:p><w:r><w:rPr><w:b/></w:rPr>'
        "<w:t>Click here.</w:t></w:r></w:p></w:sdtContent></w:sdt>",
        location="End",
    )
    empty = body.paragraphs[-2]
    assert (empty.font.italic, empty.font.bold) == (True, False)
    assert (empty.get_font(direct=True).italic, empty.get_font(direct=True).bold) == (True, False)
    shown = body.paragraphs[-1]
    assert shown.text == ""
    # Word's font reads a placeholder's runs though its text doesn't (core-ts check 52 case 13, after Accept All)
    assert shown.font.bold is True


def test_a_paragraphs_font_reads_the_runs_of_its_content_range_so_an_empty_run_at_either_end_is_not_read(new_package):
    """core-ts ``getFont`` reads ``getRange('Content').runs``: an empty run at the start or the end holds none of the text."""
    body = new_package.body
    body.insert_xml(
        '<w:p><w:r><w:t>abc</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t></w:t></w:r></w:p>'
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t></w:t></w:r><w:r><w:t>abc</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>ab</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t></w:t></w:r><w:r><w:t>c</w:t></w:r></w:p>',
        location="End",
    )
    for p, bold in zip(body.paragraphs[-3:], (False, False, None)):
        assert p.font.bold is bold
        assert p.get_range("Content").font.bold is bold


def test_a_font_writes_change_report_records_every_runs_direct_value_before_the_write_none_where_they_differ(new_package):
    p = new_package.body.insert_paragraph("plain bold plain")
    p.search("bold")[0].font.bold = True
    p.font.italic = True
    report = new_package.last_change
    assert report.operation == "format"
    assert (report.text_before, report.text_after) == ("italic=False", "italic=True")
    p.font.bold = False
    report = new_package.last_change
    # the runs' own w:b differ (CR-016 Phase B): None, not the first run's False
    assert (report.text_before, report.text_after) == ("bold=None", "bold=False")


def test_the_getters_are_typed_bool_or_none_and_float_or_none():
    import typing

    from docx4j_py.model.content import Font

    for name in ("bold", "italic", "strike_through", "double_strike_through", "subscript", "superscript"):
        assert typing.get_type_hints(getattr(Font, name).fget)["return"] == bool | None, name
    assert typing.get_type_hints(Font.size.fget)["return"] == float | None
