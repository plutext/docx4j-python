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
    # docx4j ships twenty-nine formats and CR-002 Phase B ported all of them
    # (section 15.2), so the fall-back needs a w:numFmt no library registers ---
    # one of the counting styles docx4j leaves to ``register`` (Japanese here)
    assert format_value("japaneseCounting", 5) == "5"


@pytest.mark.parametrize(
    ("num_fmt", "value", "expected"),
    [
        ("ordinal", 21, "21st"),
        ("ordinal", 12, "12th"),
        ("cardinalText", 1234, "One Thousand Two Hundred Thirty-Four"),
        ("ordinalText", 21, "Twenty-First"),
        ("ordinalText", 100, "One Hundredth"),
        ("hex", 255, "FF"),
        ("chicago", 1, "*"),
        ("chicago", 5, "**"),
        ("numberInDash", 3, "- 3 -"),
        ("decimalFullWidth", 12, "\uff11\uff12"),
        ("decimalFullWidth2", 12, "\uff11\uff12"),
        ("thaiNumbers", 12, "\u0e51\u0e52"),
        ("hindiNumbers", 12, "\u0967\u0968"),
        ("russianLower", 1, "\u0430"),
        ("russianUpper", 29, "\u0410\u0410"),
        ("arabicAlpha", 1, "\u0623"),
        ("thaiLetters", 1, "\u0e01"),
        ("hebrew1", 15, "\u05d8\u05d5"),
        ("hebrew1", 16, "\u05d8\u05d6"),
        ("hebrew1", 115, "\u05e7\u05d8\u05d5"),
        ("hebrew1", 401, "\u05ea\u05d0"),
        ("chineseCounting", 10086, "\u4e00\u4e07\u516b\u5341\u516d"),
        ("chineseCountingThousand", 21, "\u4e8c\u5341\u4e00"),
        ("chineseLegalSimplified", 21, "\u8d30\u62fe\u58f9"),
        ("decimalEnclosedCircle", 1, "\u2460"),
        ("decimalEnclosedCircle", 20, "\u2473"),
        ("decimalEnclosedCircleChinese", 3, "\u2462"),
    ],
)
def test_the_twenty_docx4j_also_ships(num_fmt, value, expected):
    """CR-002 section 15.2: every ``NumberFormat*`` class, with docx4j's values."""
    assert format_value(num_fmt, value) == expected


def test_the_exotic_formats_fall_back_where_they_cannot_say_the_value():
    assert format_value("decimalEnclosedCircle", 21) == "21"
    assert format_value("chicago", 0) == "0"
    assert format_value("hebrew1", 0) == "0"
    assert format_value("cardinalText", -1) == "-1"


def test_the_registry_is_docx4js_twenty_nine():
    from docx4j_py.model.listnumbering.formats import registered_formats

    assert len(registered_formats()) == 29


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
    """docx4j's reason strings at VERSION_17_1_1 7fba7a150, verbatim (CR-003 section 18.10)."""
    package = numbered_package(DECIMAL)
    assert Emulator.get_number_of(package, None, "99", "0") is None
    emulator = emulator_of(package)
    dangling = emulator.resolve(None, "99", "0", True)
    assert (dangling.not_numbered, dangling.reason) == (True, "no w:num for numId 99 (the paragraph's own)")
    missing = emulator.resolve(None, "7", "4", True)
    assert (missing.not_numbered, missing.reason) == (True, "no w:lvl 4 in w:num 7 (the paragraph's own)")
    styled = numbered_package(
        DECIMAL,
        '<w:style w:type="paragraph" w:styleId="L"><w:name w:val="L"/>'
        '<w:pPr><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr></w:style>',
    )
    via_style = emulator_of(styled).resolve("L", None, None, False)
    assert via_style.reason == "no w:num for numId 99 (from style 'L')"


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

    # since CR-002 Phase B the result's w:ind is the typed one docx4j hands out
    assert result.ind.left == 720
    assert result.ind.hanging == 360
    assert result.ind_resolved == result.ind, "the level states its own, so getInd is it"
    assert result.indent.to_dict() == {"left": 720, "hanging": 360}


def test_reading_the_labels_leaves_the_numbering_part_untouched():
    """A read never unmarshals: CR-003 section 13's promise, kept by the emulator."""
    package = fixture("lists.docx")
    labels(package.body)

    assert not package.numbering_definitions_part.is_unmarshalled
    assert not package.style_definitions_part.is_unmarshalled


# ---------------------------------------------------------------------------
# two readings aligned with docx4j and its TypeScript port on 2026-09-19,
# outside every fixture (CR-003 section 18.10)
# ---------------------------------------------------------------------------

LEGAL = (
    '<w:abstractNum w:abstractNumId="7">'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>'
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:isLgl/>'
    '<w:lvlText w:val="%1.%2"/></w:lvl>'
    "</w:abstractNum>"
    '<w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>'
    '<w:num w:numId="8"><w:abstractNumId w:val="7"/>'
    '<w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
    '<w:lvlText w:val="%1.%2"/></w:lvl></w:lvlOverride></w:num>'
)


def test_an_override_level_without_isLgl_turns_legal_numbering_off():
    """docx4j reads w:isLgl off the controlling level, the override when there is one."""
    package = numbered_package(LEGAL)
    state = NumberingState()
    Emulator.get_number_of(package, None, "7", "0", state=state)
    assert Emulator.get_number_of(package, None, "7", "1", state=state).num_string == "1.1", (
        "legal: the inherited level in decimal whatever its own format"
    )
    state = NumberingState()
    Emulator.get_number_of(package, None, "8", "0", state=state)
    assert Emulator.get_number_of(package, None, "8", "1", state=state).num_string == "I.1", (
        "the override level states no w:isLgl, so legal numbering is off for w:num 8"
    )


@pytest.mark.parametrize(
    ("level_text", "expected"),
    [
        ("%1.", "1."),
        ("%1%", "1"),  # a trailing % contributes nothing
        ("%1%x", "1x"),  # a % before a non-digit writes the character alone
        ("%1.%9", "1."),  # a level the definition lacks contributes nothing
        ("%", ""),
    ],
)
def test_a_malformed_lvlText_never_raises_and_reads_as_the_typescript_port_reads_it(level_text, expected):
    numbering = (
        '<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0">'
        f'<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="{level_text}"/>'
        "</w:lvl></w:abstractNum>"
        '<w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>'
    )
    package = numbered_package(numbering)
    assert Emulator.get_number_of(package, None, "7", "0").num_string == expected
