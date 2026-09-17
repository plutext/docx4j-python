"""The numbering emulator against Word. CR-002 section 6.2, CR-003 Phase H.

Every label here is one **Word** painted: the fixtures are the probe documents
of docx4j's CR-014 and CR-015 (``tests/fixtures/numbering-*.docx`` and
``styles-numpr-ilvl-only.docx``, copied from ``docx4j-layout-fidelity``'s
corpus), and the expected strings are the answers Word gave in the goldens of
2026-09-12, recorded in docx4j's
``docs/developer/change-requests/CR-014-list-numbering-model.md``.
"""

from __future__ import annotations

import pytest
from conftest import FIXTURES, fixture

from docx4j_py import create_package, load
from docx4j_py.model.listnumbering import (
    Emulator,
    NumberingState,
    emulator_of,
    format_value,
    labels_for,
)


def labels(body) -> list[str]:
    """The label of every list paragraph of a body, in document order."""
    found = labels_for(body)
    return [
        found[id(paragraph.element)].result.num_string
        for paragraph in body.iter_paragraphs()
        if id(paragraph.element) in found
    ]


def probe(name: str):
    """One of the Word-measured probe documents."""
    return load(FIXTURES / f"{name}.docx")


# ---------------------------------------------------------------------------
# the number formats (docx4j NumberFormatter)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("num_fmt", "value", "expected"),
    [
        ("decimal", 7, "7"),
        ("decimalZero", 7, "07"),
        ("decimalZero", 17, "17"),
        ("lowerLetter", 1, "a"),
        ("lowerLetter", 27, "aa"),
        ("upperLetter", 26, "Z"),
        ("lowerRoman", 4, "iv"),
        ("upperRoman", 1984, "MCMLXXXIV"),
        ("none", 3, ""),
        ("bullet", 3, "*"),
    ],
)
def test_the_eight_formats_are_docx4js(num_fmt, value, expected):
    assert format_value(num_fmt, value) == expected


def test_a_format_that_cannot_say_the_value_falls_back_to_decimal():
    # docx4j's rule: a roman numeral above 3999 is a decimal label and one
    # logged warning, never an exception that ends the export
    assert format_value("upperRoman", 4000) == "4000"
    assert format_value("lowerLetter", 0) == "0"


def test_an_unknown_numFmt_falls_back_to_decimal(caplog):
    assert format_value("chineseCounting", 5) == "5"


# ---------------------------------------------------------------------------
# the probes, and the labels Word painted (docx4j CR-014, goldens 2026-09-12)
# ---------------------------------------------------------------------------


def test_every_w_num_over_one_abstract_definition_continues_one_sequence():
    """P1 ``numbering-shared-abstract``: Word 1 2 3 4 5 6, interleaved A A B B A B."""
    assert labels(probe("numbering-shared-abstract").body) == [
        "1.",
        "2.",
        "3.",
        "4.",
        "5.",
        "6.",
    ]


def test_a_numStyleLink_definition_is_a_list_of_its_own():
    """P2 ``numbering-numstylelink-separate``: Word Y 1 2 3, X 1 2 3, Y 4 5 6."""
    assert labels(probe("numbering-numstylelink-separate").body) == [
        "1.",
        "2.",
        "3.",  # w:num 20, the numbering style's own list
        "1.",
        "2.",
        "3.",  # w:num 21, through w:numStyleLink: its own counter
        "4.",
        "5.",
        "6.",  # w:num 20 again, carrying on
    ]


def test_the_default_paragraph_style_may_be_numbered():
    """P6 ``numbering-default-style-numbered``: Word numbers the three 1 2 3.

    The paragraphs name no style at all; the ``w:default="1"`` one carries the
    ``w:numPr``. docx4j's ``getNumber(pkg, pPr)`` assumed it could not, which
    CR-014 measured as wrong.
    """
    package = probe("numbering-default-style-numbered")
    assert labels(package.body) == ["1.", "2.", "3."]

    # and the control, a style based on nothing, is not numbered
    control = [p for p in package.body.paragraphs if "Unlinked control" in p.text]
    assert control and not control[0].is_list_item


def test_w_lvlRestart_says_which_shallower_level_restarts_a_level():
    """P8 ``numbering-lvlrestart``: list A ``w:val=0``, list B ``w:val=1``."""
    both = labels(probe("numbering-lvlrestart").body)
    assert both[:8] == [
        "1.",
        "1.1.",
        "1.1.1.",
        "1.1.2.",
        "1.2.",
        "1.2.3.",  # w:val 0: nothing restarts level 2
        "2.",
        "2.1.4.",  # and a level reset but unused shows its start, not start - 1
    ]
    assert both[8:] == [
        "1.",
        "1.1.",
        "1.1.1.",
        "1.1.2.",
        "1.2.",
        "1.2.3.",  # w:val 1: level 1 does not restart it
        "2.",
        "2.1.1.",  # but level 0 does
    ]


def test_a_style_may_state_the_ilvl_and_inherit_the_numId():
    """P4 of CR-015, ``styles-numpr-ilvl-only``: Word 1. 1.1. 2. 2.1. 2.2."""
    assert labels(probe("styles-numpr-ilvl-only").body) == ["1.", "1.1.", "2.", "2.1.", "2.2."]


def test_each_story_counts_on_its_own():
    """P7 ``numbering-stories``: Word's map of the stories, measured."""
    package = probe("numbering-stories")
    main = package.main_document_part

    assert labels(package.body) == ["1.", "2.", "3.", "4.", "5.", "6."], "the body runs past the notes"
    headers = [labels(part.body) for part in main.header_parts()]
    footers = [labels(part.body) for part in main.footer_parts()]
    assert headers == [["1.", "2.", "3."]]
    assert footers == [["4.", "5.", "6."]], "a header and its footer are one story"
    assert labels(main.footnotes_part.body) == ["1.", "2.", "3.", "4.", "5.", "6."], (
        "the second footnote continues the first's count"
    )
    assert labels(main.endnotes_part.body) == ["1.", "2.", "3."]
    assert labels(main.comments_part.body) == ["1.", "2.", "3."]


def test_the_four_items_of_the_lists_fixture_are_one_list():
    """``tests/fixtures/lists.docx``: four items of one ``w:numId``, 1 2 3 4."""
    assert labels(fixture("lists.docx").body) == ["1.", "2.", "3.", "4."]


# ---------------------------------------------------------------------------
# resolution without PropertyResolver (CR-002 Phase B)
# ---------------------------------------------------------------------------


def numbered_package(numbering_xml: str, styles_xml: str | None = None):
    """A created document whose numbering (and styles) are the given XML."""
    package = create_package()
    package.id_seed = 20260917
    package.body.insert_paragraph("x")  # so that there is a main part to relate to
    from docx4j_py.model.content.lists import numbering_part_of

    part = numbering_part_of(package, create=True)
    part.set_bytes(
        (
            '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + numbering_xml
            + "</w:numbering>"
        ).encode("utf-8")
    )
    if styles_xml is not None:
        package.style_definitions_part.set_bytes(
            (
                '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                + styles_xml
                + "</w:styles>"
            ).encode("utf-8")
        )
    from docx4j_py.model.listnumbering import invalidate

    invalidate(package)
    return package


DECIMAL = (
    '<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0">'
    '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
    '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
    '<w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>'
)


def test_w_numId_zero_turns_numbering_off():
    package = numbered_package(DECIMAL)
    emulator = emulator_of(package)

    assert emulator.resolve(None, "7", "0", True).num_id == "7"
    off = emulator.resolve(None, "0", "0", True)
    assert off.not_numbered and "turns numbering off" in off.reason


def test_a_numId_the_document_does_not_define_is_not_numbered():
    package = numbered_package(DECIMAL)
    assert Emulator.get_number_of(package, None, "99", "0") is None


def test_a_level_linked_to_another_paragraph_style_numbers_only_that_style():
    """docx4j ``styleLinkedElsewhere``, measured on probe ``numbering-label-ilvl0``."""
    numbering = (
        '<w:abstractNum w:abstractNumId="8"><w:lvl w:ilvl="0">'
        '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="NumLinked"/>'
        '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        '<w:num w:numId="8"><w:abstractNumId w:val="8"/></w:num>'
    )
    styles = (
        '<w:style w:type="paragraph" w:styleId="NumLinked"><w:name w:val="Num Linked"/>'
        '<w:pPr><w:numPr><w:numId w:val="8"/></w:numPr></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Other"><w:name w:val="Other"/>'
        '<w:pPr><w:numPr><w:numId w:val="8"/></w:numPr></w:pPr></w:style>'
    )
    package = numbered_package(numbering, styles)

    # the style the level names is numbered
    assert Emulator.get_number_of(package, "NumLinked", None, None).num_string == "1."
    # a second style carrying the same w:numPr is not
    assert Emulator.get_number_of(package, "Other", None, None) is None
    # but direct formatting always applies
    assert Emulator.get_number_of(package, "Other", "8", "0", True).num_string == "2."


def test_peek_does_not_take_the_number():
    package = numbered_package(DECIMAL)
    emulator = emulator_of(package)
    state = NumberingState()

    assert emulator.number_of(None, "7", "0", True, state).num_string == "1."
    assert emulator.peek_number(None, state) is None  # no w:pPr at all
    assert emulator.number_of(None, "7", "0", True, state.copy()).num_string == "2."
    assert emulator.number_of(None, "7", "0", True, state).num_string == "2.", (
        "the copy the peek counted in did not touch this one"
    )


def test_the_level_indent_comes_through_on_the_result():
    package = numbered_package(DECIMAL)
    result = Emulator.get_number_of(package, None, "7", "0")

    assert result.ind.left == 720
    assert result.ind.hanging == 360
    assert result.ind.to_dict() == {"left": 720, "hanging": 360}


def test_reading_the_labels_leaves_the_numbering_part_untouched():
    """A read never unmarshals: CR-003 section 13's promise, kept by the emulator."""
    package = fixture("lists.docx")
    labels(package.body)

    assert not package.numbering_definitions_part.is_unmarshalled
    assert not package.style_definitions_part.is_unmarshalled
