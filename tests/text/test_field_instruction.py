"""``FieldInstruction``, the one parser of a field instruction. CR-025 Phase B, D6.

docx4j's ``FieldInstructionTest`` at ``c3e9920e5`` (its CR-067 D2): the name, the
parameters as Word tokenises them, the argument and the switches; and that it
reads every instruction exactly as the parser it replaces did. Here the
replaced parser is this repository's own ``FldSimpleModel`` at ``3c63804``,
before CR-025: :data:`SHAPES` holds what that class read for each of docx4j's
31 shapes (its name, its ``parameter_string``, its parameters, and the
``\\*`` and ``\\@`` switch values ``find_first_switch_value`` found in them),
computed from that commit's source, and both classes are held to it.
"""

from __future__ import annotations

import pytest

from docx4j_py.model.fields import (
    FieldInstruction,
    FldSimpleModel,
    find_first_switch_value,
    split_parameters,
    unquote,
)

#: docx4j's shapes, and what ``FldSimpleModel`` read for each at ``3c63804``:
#: (instruction, name, parameter_string, parameters, the ``\*`` value, the ``\@`` value).
SHAPES = [
    (' DATE \\@ "d MMMM yyyy" ', 'DATE', '\\@ "d MMMM yyyy"', ['\\@', '"d MMMM yyyy"'], None, 'd MMMM yyyy'),
    ('DATE', 'DATE', None, [], None, None),
    (' TIME \\@ "HH:mm" \\* MERGEFORMAT ', 'TIME', '\\@ "HH:mm" \\* MERGEFORMAT', ['\\@', '"HH:mm"', '\\*', 'MERGEFORMAT'], '', 'HH:mm'),
    (' MERGEFIELD "First Name" \\* Upper ', 'MERGEFIELD', '"First Name" \\* Upper', ['"First Name"', '\\*', 'Upper'], 'Upper', None),
    (' MERGEFIELD name \\* MERGEFORMAT ', 'MERGEFIELD', 'name \\* MERGEFORMAT', ['name', '\\*', 'MERGEFORMAT'], '', None),
    ('MERGEFIELD  Nachname ', 'MERGEFIELD', 'Nachname', ['Nachname'], None, None),
    (' DOCPROPERTY "My Prop" \\* CharFormat ', 'DOCPROPERTY', '"My Prop" \\* CharFormat', ['"My Prop"', '\\*', 'CharFormat'], 'CharFormat', None),
    ('DOCPROPERTY client', 'DOCPROPERTY', 'client', ['client'], None, None),
    (' DOCVARIABLE v1 ', 'DOCVARIABLE', 'v1', ['v1'], None, None),
    ('=SUM(ABOVE) \\# "#,##0.00"', '=', 'SUM(ABOVE) \\# "#,##0.00"', ['SUM(ABOVE)', '\\#', '"#,##0.00"'], None, None),
    ('= 1 + 2', '=', '1 + 2', ['1', '+', '2'], None, None),
    (' PAGEREF _Toc123456 \\h ', 'PAGEREF', '_Toc123456 \\h', ['_Toc123456', '\\h'], None, None),
    (' REF bookmark1 \\h \\* MERGEFORMAT ', 'REF', 'bookmark1 \\h \\* MERGEFORMAT', ['bookmark1', '\\h', '\\*', 'MERGEFORMAT'], '', None),
    (' SEQ Figure \\* ARABIC ', 'SEQ', 'Figure \\* ARABIC', ['Figure', '\\*', 'ARABIC'], 'ARABIC', None),
    (' SEQ Table \\* ARABIC \\r 3 ', 'SEQ', 'Table \\* ARABIC \\r 3', ['Table', '\\*', 'ARABIC', '\\r', '3'], 'ARABIC', None),
    (' SEQ Figure \\c ', 'SEQ', 'Figure \\c', ['Figure', '\\c'], None, None),
    (' HYPERLINK "http://example.com/x" \\l "anchor" \\o "a tip" ', 'HYPERLINK', '"http://example.com/x" \\l "anchor" \\o "a tip"', ['"http://example.com/x"', '\\l', '"anchor"', '\\o', '"a tip"'], None, None),
    (' HYPERLINK \\l "_Toc1" ', 'HYPERLINK', '\\l "_Toc1"', ['\\l', '"_Toc1"'], None, None),
    (' PAGE \\* MERGEFORMAT ', 'PAGE', '\\* MERGEFORMAT', ['\\*', 'MERGEFORMAT'], '', None),
    (' NUMPAGES ', 'NUMPAGES', '', [], None, None),
    (' TOC \\o "1-3" \\h \\z \\u ', 'TOC', '\\o "1-3" \\h \\z \\u', ['\\o', '"1-3"', '\\h', '\\z', '\\u'], None, None),
    (' STYLEREF 1 \\s ', 'STYLEREF', '1 \\s', ['1', '\\s'], None, None),
    (' AUTHOR "Jo Bloggs" ', 'AUTHOR', '"Jo Bloggs"', ['"Jo Bloggs"'], None, None),
    (' TITLE ', 'TITLE', '', [], None, None),
    (' QUOTE "hello world" ', 'QUOTE', '"hello world"', ['"hello world"'], None, None),
    (' IF { MERGEFIELD x } = "1" "yes" "no" ', 'IF', '{ MERGEFIELD x } = "1" "yes" "no"', ['{', 'MERGEFIELD', 'x', '}', '=', '"1"', '"yes"', '"no"'], None, None),
    (' FORMTEXT ', 'FORMTEXT', '', [], None, None),
    (' INCLUDEPICTURE "C:\\\\pic.png" \\d ', 'INCLUDEPICTURE', '"C:\\\\pic.png" \\d', ['"C:\\\\pic.png"', '\\d'], None, None),
    ('', None, None, [], None, None),
    ('   ', None, None, [], None, None),
    ('\\* MERGEFORMAT', '\\*', 'MERGEFORMAT', ['MERGEFORMAT'], None, None),
]


def test_the_31_shapes():
    assert len(SHAPES) == 31


@pytest.mark.parametrize("cls", [FieldInstruction, FldSimpleModel])
@pytest.mark.parametrize(("instr", "name", "parameter_string", "parameters", "general", "date"), SHAPES)
def test_every_shape_reads_as_fldsimplemodel_read_it(cls, instr, name, parameter_string, parameters, general, date):
    """docx4j ``readsAsFldSimpleModelRead``: every accessor a consumer read is as it was."""
    parsed = cls(instr)
    assert parsed.name == name
    assert parsed.parameter_text == parameter_string
    assert parsed.parameters == parameters
    assert find_first_switch_value("\\*", parsed.parameters, True) == general
    assert find_first_switch_value("\\@", parsed.parameters, True) == date
    assert parsed.switch_value("\\*") == general, "FieldInstruction's own reading of the general switch"
    assert parsed.switch_value("\\@") == date
    assert parsed.has_switch("\\h") == ("\\h" in parameters)
    if cls is FldSimpleModel:
        assert parsed.parameter_string == parameter_string
    assert FieldInstruction.parse(instr).parameters == parameters


def test_name_argument_and_switches():
    i = FieldInstruction(' DATE \\@ "d MMMM yyyy" ')
    assert i.name == "DATE"
    assert i.parameter_text == '\\@ "d MMMM yyyy"'
    assert i.parameters == ["\\@", '"d MMMM yyyy"']
    assert i.argument is None, "a switch is not the argument"
    assert i.has_switch("\\@")
    assert i.switch_value("\\@") == "d MMMM yyyy"
    assert i.switch_value("\\#") is None

    i = FieldInstruction(' MERGEFIELD "First Name" \\* Upper ')
    assert i.name == "MERGEFIELD"
    assert i.argument == "First Name", "the quotes come off the argument"
    assert i.parameters[0] == '"First Name"'
    assert i.switch_value("\\*") == "Upper"

    i = FieldInstruction("mergefield name \\* MERGEFORMAT")
    assert i.name == "MERGEFIELD", "the name is upper-cased"
    assert i.has_switch("\\*")
    assert i.switch_value("\\*") == "", "MERGEFORMAT is not a format: the switch is present with no value"
    assert i.switch_values("\\*") == ["MERGEFORMAT"]

    i = FieldInstruction(" SEQ Table \\* ARABIC \\r 3 ")
    assert i.argument == "Table"
    assert i.switch_value("\\*") == "ARABIC"
    assert i.switch_value("\\r") == "3"
    assert i.index_of_switch("\\*", 0) == 1
    assert i.index_of_switch("\\*", 2) == -1
    assert i.value_at(2) == "ARABIC"
    assert i.value_at(3) is None, "a switch is never a value"
    assert i.value_at(99) is None

    i = FieldInstruction(" PAGEREF _Toc123456 \\h ")
    assert i.argument == "_Toc123456"
    assert i.has_switch("\\h")
    assert i.switch_value("\\h") == "", "a switch with no value"
    assert not i.has_switch("\\H")


def test_a_formula_is_named_by_its_equals_sign():
    i = FieldInstruction('=SUM(ABOVE) \\# "#,##0.00"')
    assert i.name == "="
    assert i.parameter_text == 'SUM(ABOVE) \\# "#,##0.00"'
    assert i.argument == "SUM(ABOVE)"
    assert i.switch_value("\\#") == "#,##0.00"
    assert FieldInstruction("= 1 + 2").name == "="


def test_quoted_parameters_keep_their_quotes_and_their_spaces():
    i = FieldInstruction(' HYPERLINK "http://example.com/x" \\l "anchor" \\o "a tip" ')
    assert i.parameters == ['"http://example.com/x"', "\\l", '"anchor"', "\\o", '"a tip"']
    assert i.argument == "http://example.com/x"
    assert i.switch_value("\\l") == "anchor"
    assert i.switch_value("\\o") == "a tip"
    assert split_parameters('"a b"  c') == ['"a b"', "c"]
    assert unquote('"a b"') == "a b"
    assert unquote('"') == '"'
    assert unquote(None) is None


@pytest.mark.parametrize("instr", [None, "", "   "])
def test_empty_and_blank_instructions(instr):
    i = FieldInstruction(instr)
    assert i.name is None
    assert i.parameter_text is None
    assert i.parameters == []
    assert i.argument is None
    assert not i.has_switch("\\*")


def test_the_first_token_is_the_name_whatever_it_is():
    assert FieldInstruction("\\* MERGEFORMAT").name == "\\*"


def test_the_parameters_are_unmodifiable():
    """docx4j's are an unmodifiable list; here :attr:`parameters` is a copy."""
    i = FieldInstruction(" SEQ Figure ")
    i.parameters.append("x")
    assert i.parameters == ["Figure"]
