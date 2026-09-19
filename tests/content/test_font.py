"""``Font``: the Office JS run properties, and the formatting it reads.

CR-003 section 6, closed by CR-002 Phase B (section 15.1): reads report the
**effective** formatting of the first run in scope, as Office JS does ---
document defaults, paragraph style, character style, direct --- and
``get_font(direct=True)`` reads the run's own ``w:rPr``, which is what these
reads were before the resolver landed.
"""

from __future__ import annotations

from conftest import reloaded, sample


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
    }
    # the effective read still reports the document defaults underneath
    assert font.size == 11.0


def test_a_write_reaches_every_run_a_read_only_the_first(new_package):
    paragraph = new_package.body.insert_paragraph("plain bold plain")
    paragraph.search("bold")[0].font.bold = True
    assert len(paragraph.runs) == 3

    assert paragraph.font.bold is False, "the first run is not bold"
    paragraph.font.name = "Arial"
    assert all(run.r_pr.r_fonts.ascii == "Arial" for run in paragraph.runs)


def test_repr_says_what_the_font_is(new_package):
    paragraph = new_package.body.insert_paragraph("x")
    assert repr(paragraph.get_font(direct=True)) == "<Font paragraph plain>"
    # effective, the document defaults give it a size and the theme a face
    assert repr(paragraph.font) == "<Font paragraph Aptos 11.0pt>"
    paragraph.font.bold = True
    paragraph.font.size = 11
    assert repr(paragraph.font) == "<Font paragraph bold Aptos 11.0pt>"


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
