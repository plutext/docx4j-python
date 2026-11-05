"""``w:numId`` 0 clears every inherited left, hanging and first-line indent (CR-036).

docx4j ``9e3100351``'s rule, from Word's goldens of 5 November 2026 on its probe
``numbering-off-style-ind`` (``docx4j-export-fo/docs/word-layout-rules.md``, anchor
``s28numid0``). ``CASES`` are docx4j ``NumIdZeroIndentTest``'s, read here as the resolver
answers them, in twips. The level is left 1440, hanging 1152.
"""

from __future__ import annotations

import pytest
from docx4j_xsdata.exceptions import ConverterWarning

from docx4j_py import create_package
from docx4j_py.model.content.lists import numbering_part_of
from docx4j_py.model.text import extract

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

NUMBERING = f"""<w:numbering {W}>
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="1152"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="500" w:right="57"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>"""


def style(style_id: str, p_pr: str, based_on: str | None = None) -> str:
    base = f'<w:basedOn w:val="{based_on}"/>' if based_on else ""
    return f'<w:style w:type="paragraph" w:styleId="{style_id}"><w:name w:val="{style_id}"/>{base}<w:pPr>{p_pr}</w:pPr></w:style>'


def num_pr(num_id: str) -> str:
    return f'<w:numPr><w:numId w:val="{num_id}"/></w:numPr>'


STYLES = (
    f"<w:styles {W}>"
    + "".join(
        (
            style("Numbered", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
            # a numbered style stating its own w:ind, as docx4j's corpus document 4083's 2_Question
            style("OwnInd", num_pr("1") + '<w:ind w:left="454" w:right="57" w:hanging="454"/>'),
            # a w:ind on the style a numbered one is based on
            style("IndBase", '<w:ind w:left="720" w:hanging="360"/>'),
            style("OnBase", num_pr("1"), "IndBase"),
            # w:numId 0 in a style, and with a w:ind of its own
            style("Switch", num_pr("0"), "OwnInd"),
            style("SwitchInd", num_pr("0") + '<w:ind w:left="300" w:hanging="100"/>', "OwnInd"),
            # the rest are not docx4j's cases
            style("NoList", num_pr("9") + '<w:ind w:left="454" w:hanging="454"/>'),
            style("NoLevelInd", num_pr("2") + '<w:ind w:left="454" w:hanging="454"/>'),
            style("RightAsLevel", num_pr("3")),
            style("Plain", '<w:ind w:left="454" w:hanging="454"/>'),
            style("FirstLine", num_pr("1") + '<w:ind w:left="454" w:firstLine="227"/>'),
            style(
                "Forms",
                num_pr("1")
                + '<w:ind w:start="454" w:startChars="100" w:leftChars="100" w:end="57" w:endChars="9"'
                ' w:rightChars="9" w:hangingChars="50" w:firstLineChars="50"/>',
            ),
        )
    )
    + "</w:styles>"
)

OFF = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'


def paragraph(style_id: str, p_pr: str = "", *, numbering: bool = True, styles: str = STYLES):
    """A package whose one paragraph has this style and these properties, and that paragraph's view."""
    package = create_package()
    package.body.insert_paragraph("heading")
    if numbering:
        numbering_part_of(package, create=True).set_bytes(NUMBERING.encode("utf-8"))
    package.style_definitions_part.set_bytes(styles.encode("utf-8"))
    document = package.main_document_part
    document.set_bytes(
        f'<w:document {W}><w:body><w:p><w:pPr><w:pStyle w:val="{style_id}"/>{p_pr}</w:pPr>'
        f"<w:r><w:t>heading</w:t></w:r></w:p></w:body></w:document>".encode()
    )
    return package, package.body.paragraphs[0]


def indent(style_id: str, p_pr: str = "", **options) -> tuple:
    """The effective indent as (left, hanging, first line, right), None for each it doesn't state."""
    _package, view = paragraph(style_id, p_pr, **options)
    ind = view.effective_p_pr.ind
    return (
        (None, None, None, None)
        if ind is None
        else (ind.left, ind.hanging, ind.first_line, ind.right)
    )


CASES = {
    "a numbered paragraph: the level's": ("Numbered", "", (1440, 1152, None, None)),
    "the paragraph's own w:ind alone": (
        "Numbered",
        OFF + '<w:ind w:left="720"/>',
        (720, None, None, None),
    ),
    "N1, the style's own left and hanging go, its right stays": (
        "OwnInd",
        OFF,
        (None, None, None, 57),
    ),
    "N4, and the w:ind of the style it is based on": ("OnBase", OFF, (None, None, None, None)),
    "N7, w:numId 0 in a style": ("Switch", "", (None, None, None, 57)),
    "N8, a w:ind in the style that switches off stands": ("SwitchInd", "", (300, 100, None, 57)),
    "N10, a direct left, no hanging": (
        "OwnInd",
        OFF + '<w:ind w:left="300"/>',
        (300, None, None, 57),
    ),
    "N11, a direct hanging, no left": (
        "OwnInd",
        OFF + '<w:ind w:hanging="200"/>',
        (None, 200, None, 57),
    ),
    "N14, w:numId 0 in a style, a direct left": (
        "Switch",
        '<w:ind w:left="142"/>',
        (142, None, None, 57),
    ),
}


@pytest.mark.parametrize(("style_id", "p_pr", "expected"), CASES.values(), ids=CASES.keys())
def test_docx4js_cases(style_id, p_pr, expected):
    assert indent(style_id, p_pr) == expected


def test_a_first_line_indent_goes_as_a_hanging_one_does():
    assert indent("FirstLine", OFF) == (None, None, None, None)
    assert indent("FirstLine", OFF + '<w:ind w:firstLine="100"/>') == (None, None, 100, None)


def test_a_style_asked_by_name_reads_the_same():
    package, _view = paragraph("Switch")
    ind = package.get_property_resolver().get_effective_ppr("Switch").ind
    assert (ind.left, ind.hanging, ind.right) == (None, None, 57)


def test_every_left_and_first_line_form_clears_and_every_right_form_stays():
    _package, view = paragraph("Forms", OFF)
    ind = view.effective_p_pr.ind
    gone = (
        "left",
        "left_chars",
        "start",
        "start_chars",
        "hanging",
        "hanging_chars",
        "first_line",
        "first_line_chars",
    )
    assert {name: getattr(ind, name) for name in gone} == dict.fromkeys(gone)
    assert (ind.end, ind.end_chars, ind.right_chars) == (57, 9, 9)


def test_a_right_indent_equal_to_the_levels_stays():
    """The old rule took out what equalled the level's, the right indent with it."""
    assert indent("RightAsLevel", OFF) == (None, None, None, 57)


def test_a_w_num_id_naming_no_list_leaves_the_indent_alone():
    assert indent("NoList", OFF) == (454, 454, None, None)


def test_no_numbering_part_leaves_the_indent_alone():
    assert indent("OwnInd", OFF, numbering=False) == (454, 454, None, 57)


def test_a_level_with_no_w_ind_clears_as_docx4js_does():
    """Not measured: the rule's document leaves this shape out, and docx4j's code clears it."""
    assert indent("NoLevelInd", OFF) == (None, None, None, None)


def test_a_w_num_id_is_read_as_word_reads_it():
    """``"1.0"`` is list 1 (CR-021), as ``get_ind`` reads it."""
    styles = STYLES.replace(
        "</w:styles>", style("Decimal", num_pr("1.0") + '<w:ind w:left="454"/>') + "</w:styles>"
    )
    with pytest.warns(ConverterWarning):
        assert indent("Decimal", OFF, styles=styles) == (None, None, None, None)


@pytest.mark.parametrize("spelling", ["0.0", "0.9"])
def test_the_w_num_id_0_itself_is_read_as_word_reads_it(spelling):
    """A decimal truncated toward zero (CR-021), as the numbering emulator reads it."""
    with pytest.warns(ConverterWarning):
        assert indent("OwnInd", num_pr(spelling)) == (None, None, None, 57)


def test_a_w_num_id_word_refuses_turns_nothing_off():
    with pytest.warns(ConverterWarning):
        assert indent("OwnInd", num_pr("abc")) == (454, 454, None, 57)


def test_a_second_w_num_id_0_clears_nothing_more():
    """Numbering is off already, so the w:ind stated with the first stands."""
    assert indent("SwitchInd", OFF) == (300, 100, None, 57)


def test_nothing_inherited_changes_nothing():
    assert indent("Plain", OFF) == (454, 454, None, None)


def test_a_read_changes_no_style():
    """The rule clears a copy: the style's own ``w:ind``, in the tree and as resolved, is as it was."""
    package, view = paragraph("OwnInd", OFF)
    assert view.effective_p_pr.ind.left is None
    styles = package.style_definitions_part.contents.style
    own = next(entry for entry in styles if entry.style_id == "OwnInd")
    assert (own.p_pr.ind.left, own.p_pr.ind.hanging) == (454, 454)
    resolved = package.get_property_resolver().get_effective_ppr("OwnInd").ind
    assert (resolved.left, resolved.hanging) == (454, 454)


def test_the_content_api_reads_it_in_points():
    _package, view = paragraph("OwnInd", OFF + '<w:ind w:left="300"/>')
    assert (view.left_indent, view.first_line_indent, view.right_indent) == (15.0, 0.0, 2.85)


def test_the_text_model_reads_it():
    package, _view = paragraph("OwnInd", OFF + '<w:ind w:left="300"/>')
    block = extract(package).blocks[0]
    assert (block.indent_left, block.indent_hanging) == (300, None)
