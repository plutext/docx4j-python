"""A field's formatting switches, ``docx4j_py.model.fields``. CR-011 Phase C (section 13).

docx4j's ``FormattingSwitchHelperNumericTests``, ``...GeneralTests`` and
``...NoSwitchTests`` at ``198d015a3``: every case they run with docx4j's
default properties, by their data (the switch, the value, the result, each
under ``MERGEFIELD`` and ``DOCPROPERTY``); a case docx4j expects to refuse
raises :class:`~docx4j_py.model.fields.FieldFormattingError` here too. The
four cases that set ``docx4j.Fields.Numbers.JavaStylePercentHandling`` are
not ported, the property being off by default. And dates in docx4j's
international inference (``FormattingSwitchHelperDateTests`` assumes its US
one), with ISO 8601 input as its ``testDateIso*``.
"""

from __future__ import annotations

import pytest

from docx4j_py.model.fields import (
    FieldFormattingError,
    FldSimpleModel,
    apply_formatting_switch,
    convert_date_pattern,
    determine_date_format,
    extract_number,
)

REFUSED = "class org.docx4j.model.fields.FieldFormattingException"

#: (docx4j's test, the fields it runs under, the switch, the value, the result or REFUSED).
CASES = [
    ('Numeric.testNumberNotSpecified', ('DOCPROPERTY', 'MERGEFIELD'), '\\# ', '123', 'Error! Switch argument not specified.'),
    ('Numeric.testNumber31', ('DOCPROPERTY', 'MERGEFIELD'), '\\# 00.00', '9', '09.00'),
    ('Numeric.testNumber32', ('DOCPROPERTY', 'MERGEFIELD'), '\\# 00.00', '9.006', '09.01'),
    ('Numeric.testNumber33', ('DOCPROPERTY', 'MERGEFIELD'), '\\# $###', '15', '$15'),
    ('Numeric.testNumber34', ('DOCPROPERTY', 'MERGEFIELD'), '\\# ##.##', '9', '9'),
    ('Numeric.testNumber35', ('DOCPROPERTY', 'MERGEFIELD'), '\\# ##.##', '9.006', '9.01'),
    ('Numeric.testNumber36', ('DOCPROPERTY', 'MERGEFIELD'), '\\# ##.##0000', '9.006', '9.006'),
    ('Numeric.testNumber37', ('DOCPROPERTY', 'MERGEFIELD'), '\\# ##.##00#0', '9.006', '9.006'),
    ('Numeric.testNumber38', ('DOCPROPERTY', 'MERGEFIELD'), '\\# ####.0000', '123.456', '123.4560'),
    ('Numeric.testNumber39', ('DOCPROPERTY', 'MERGEFIELD'), '\\# 0000.####', '123.456', '0123.456'),
    ('Numeric.testNumber40', ('DOCPROPERTY', 'MERGEFIELD'), '\\# #.0000', '123.456', '123.4560'),
    ('Numeric.testNumber41', ('DOCPROPERTY', 'MERGEFIELD'), '\\# 0.####', '123.456', '123.456'),
    ('Numeric.testNumber44', ('DOCPROPERTY', 'MERGEFIELD'), '\\# 0.00x', '0.125678', '0.126'),
    ('Numeric.testNumber45', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .x', '0.75', '.8'),
    ('Numeric.testNumber46', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .000x', '0.75', '.7500'),
    ('Numeric.testNumber47', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .###x', '0.75', '.75'),
    ('Numeric.testNumber48', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .###x', '.75', '.75'),
    ('Numeric.testNumber55', ('DOCPROPERTY', 'MERGEFIELD'), '\\# 0.00xx', '0.125678', '0.1257'),
    ('Numeric.testNumber56', ('DOCPROPERTY', 'MERGEFIELD'), '\\# 0.x0x', '0.125678', '0.126'),
    ('Numeric.testNumber57', ('DOCPROPERTY', 'MERGEFIELD'), '\\# $###.00', '95.4', '$95.40'),
    ('Numeric.testNumber60', ('DOCPROPERTY', 'MERGEFIELD'), '\\# $#,###,###', '2456800', '$2,456,800'),
    ('Numeric.testNumber61', ('DOCPROPERTY', 'MERGEFIELD'), '\\# -##', '-10', '-10'),
    ('Numeric.testNumber62', ('DOCPROPERTY', 'MERGEFIELD'), '\\# -##', '10', '10'),
    ('Numeric.testNumber63', ('DOCPROPERTY', 'MERGEFIELD'), '\\# +##', '10', '10'),
    ('Numeric.testNumber64', ('DOCPROPERTY', 'MERGEFIELD'), '\\# +##', '-10', '-10'),
    ('Numeric.testNumber65', ('DOCPROPERTY', 'MERGEFIELD'), '\\# ##%', '33', '33%'),
    ('Numeric.testNumber68', ('DOCPROPERTY', 'MERGEFIELD'), '\\# "$##0.00 \'is the sales tax\'"', '3.89', '$3.89 is the sales tax'),
    ('Numeric.testNumber69', ('DOCPROPERTY', 'MERGEFIELD'), '\\# "$##0.00 \'is the ##.00 goodness\'"', '3.89', '$3.89 is the ##.00 goodness'),
    ('Numeric.testNumber70', ('DOCPROPERTY', 'MERGEFIELD'), '\\# "$##0.00 is the \'##.00\' goodness"', '3.89', '$3.89 is the ##.00 goodness'),
    ('Numeric.testNumber72', ('DOCPROPERTY', 'MERGEFIELD'), '\\# 00', '€180,000.00 EUR', '180000'),
    ('Numeric.testNumber42', ('DOCPROPERTY', 'MERGEFIELD'), '\\# x#.#x', '123456', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber43', ('DOCPROPERTY', 'MERGEFIELD'), '\\# x#.#x', '123.456', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber49', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .###x', '-0.75', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber50', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .###x', '-.75', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber51', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .###', '-0.75', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber52', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .###', '-.75', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber53', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .000', '-0.75', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber54', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .000', '-.75', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber58', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .00', '95.4', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber59', ('DOCPROPERTY', 'MERGEFIELD'), '\\# .##', '95.4', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber66', ('DOCPROPERTY', 'MERGEFIELD'), '\\# #%#', '33', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber67', ('DOCPROPERTY', 'MERGEFIELD'), '\\# #$#', '33', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber71', ('DOCPROPERTY', 'MERGEFIELD'), "\\# #'y'#", '34', 'class org.docx4j.model.fields.FieldFormattingException'),
    ('Numeric.testNumber1', ('DOCPROPERTY', 'MERGEFIELD'), '\\# ', 'AA', 'Error! Switch argument not specified.'),
    ('Numeric.testNumber2', ('DOCPROPERTY', 'MERGEFIELD'), '\\# ', '123.4500', 'Error! Switch argument not specified.'),
    ('General.testStringNotSpecified', ('DOCPROPERTY', 'MERGEFIELD'), '\\* ', 'mary smith', 'Error! Switch argument not specified.'),
    ('General.testStringCaps', ('DOCPROPERTY', 'MERGEFIELD'), '\\* Caps', 'mary smith', 'Mary Smith'),
    ('General.testStringUpper', ('DOCPROPERTY', 'MERGEFIELD'), '\\* Upper', 'mary smith', 'MARY SMITH'),
    ('General.testStringNone1', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', 'Mary SmiTH', 'Mary SmiTH'),
    ('General.testStringNone2', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '01', '01'),
    ('General.testStringNone3', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.1', '0.1'),
    ('General.testStringNone4', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.0', '0.0'),
    ('General.testStringNone5', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.00', '0.00'),
    ('General.testStringNone6', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.', '0.'),
    ('General.testStringNone7', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.1 A', '0.1 A'),
    ('General.testStringNone8', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.1 1', '0.1 1'),
    ('General.testStringNone9', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.1 .', '0.1 .'),
    ('General.testStringNone10', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.00 0', '0.00 0'),
    ('General.testStringNone11', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.00 1', '0.00 1'),
    ('General.testStringNone12', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0.00 A', '0.00 A'),
    ('General.testStringNone13', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0000123456', '0000123456'),
    ('General.testStringNone14', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '000012345.006', '000012345.006'),
    ('General.testStringNone15', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0000123AA456', '0000123AA456'),
    ('General.testStringNone16', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '0000123AA45.006', '0000123AA45.006'),
    ('General.testStringCapsQuote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* Caps', '"mary smith"', '"Mary Smith"'),
    ('General.testStringCaps2Quote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* Caps', '"marysmith"', '"Marysmith"'),
    ('General.testStringFirstCapQuote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* FirstCap', '"mary smith"', '"Mary smith"'),
    ('General.testStringLowerQuote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* Lower', '"Mary Smith"', '"mary smith"'),
    ('General.testStringUpperQuote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* Upper', '"Mary Smith"', '"MARY SMITH"'),
    ('General.testStringNone1Quote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '"Mary SmiTH"', '"Mary SmiTH"'),
    ('General.testStringNone2Quote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '"mary SmiTH"', '"mary SmiTH"'),
    ('General.testStringNone3Quote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '"0.1 A"', '"0.1 A"'),
    ('General.testStringNone4Quote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '"0.1 1"', '"0.1 1"'),
    ('General.testStringNone5Quote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '"0.1 ."', '"0.1 ."'),
    ('General.testStringNone6Quote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '"0.00 0"', '"0.00 0"'),
    ('General.testStringNone7Quote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '"0.00 1"', '"0.00 1"'),
    ('General.testStringNone8Quote', ('DOCPROPERTY', 'MERGEFIELD'), '\\* MERGEFORMAT', '"0.00 A"', '"0.00 A"'),
    ('NoSwitch.testNone1', ('DOCPROPERTY', 'MERGEFIELD'), '', 'mary smith', 'mary smith'),
    ('NoSwitch.testNone2', ('DOCPROPERTY', 'MERGEFIELD'), '', 'mary"smith', 'mary"smith'),
    ('NoSwitch.testNone3', ('DOCPROPERTY', 'MERGEFIELD'), '', '"mary smith"', '"mary smith"'),
    ('NoSwitch.testNone4', ('DOCPROPERTY', 'MERGEFIELD'), '', '"marysmith"', '"marysmith"'),
    ('NoSwitch.testNone5', ('DOCPROPERTY', 'MERGEFIELD'), '', '"mary smith"', '"mary smith"'),
    ('NoSwitch.testNone6', ('DOCPROPERTY', 'MERGEFIELD'), '', '"Mary Smith"', '"Mary Smith"'),
    ('NoSwitch.testNone7', ('DOCPROPERTY', 'MERGEFIELD'), '', '"Mary Smith"', '"Mary Smith"'),
    ('NoSwitch.testNone8', ('DOCPROPERTY', 'MERGEFIELD'), '', '"Mary Smith" Capoop', '"Mary Smith" Capoop'),
    ('NoSwitch.testNone9', ('DOCPROPERTY', 'MERGEFIELD'), '', '"Mary SmiTH"', '"Mary SmiTH"'),
    ('NoSwitch.testNone10', ('DOCPROPERTY', 'MERGEFIELD'), '', '"mary SmiTH"', '"mary SmiTH"'),
    ('NoSwitch.testNone11', ('DOCPROPERTY', 'MERGEFIELD'), '', 'mary smith', 'mary smith'),
    ('NoSwitch.testNone12', ('DOCPROPERTY', 'MERGEFIELD'), '', 'Mary SmiTH', 'Mary SmiTH'),
    ('NoSwitch.testNone13', ('DOCPROPERTY', 'MERGEFIELD'), '', '"mary smith"', '"mary smith"'),
    ('NoSwitch.testNone14', ('DOCPROPERTY', 'MERGEFIELD'), '', '01', '01'),
    ('NoSwitch.testNone15', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.1', '0.1'),
    ('NoSwitch.testNone16', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.0', '0.0'),
    ('NoSwitch.testNone17', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.00', '0.00'),
    ('NoSwitch.testNone18', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.', '0.'),
    ('NoSwitch.testNone19', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.1 A"', '"0.1 A"'),
    ('NoSwitch.testNone20', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.1 1"', '"0.1 1"'),
    ('NoSwitch.testNone21', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.1 ."', '"0.1 ."'),
    ('NoSwitch.testNone22', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.00 0"', '"0.00 0"'),
    ('NoSwitch.testNone23', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.00 1"', '"0.00 1"'),
    ('NoSwitch.testNone24', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.00 A"', '"0.00 A"'),
    ('NoSwitch.testNone25', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.1 A', '0.1 A'),
    ('NoSwitch.testNone26', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.1 1', '0.1 1'),
    ('NoSwitch.testNone27', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.1 .', '0.1 .'),
    ('NoSwitch.testNone28', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.00 0', '0.00 0'),
    ('NoSwitch.testNone29', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.00 1', '0.00 1'),
    ('NoSwitch.testNone30', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.00 A', '0.00 A'),
    ('NoSwitch.testNone31', ('DOCPROPERTY', 'MERGEFIELD'), '', '0000123456', '0000123456'),
    ('NoSwitch.testNone32', ('DOCPROPERTY', 'MERGEFIELD'), '', '000012345.006', '000012345.006'),
    ('NoSwitch.testNone33', ('DOCPROPERTY', 'MERGEFIELD'), '', '0000123AA456', '0000123AA456'),
    ('NoSwitch.testNone34', ('DOCPROPERTY', 'MERGEFIELD'), '', '0000123AA45.006', '0000123AA45.006'),
    ('NoSwitch.testNone35', ('DOCPROPERTY', 'MERGEFIELD'), '', '123', '123'),
    ('NoSwitch.testNone36', ('DOCPROPERTY', 'MERGEFIELD'), '', '-123', '-123'),
    ('NoSwitch.testNone37', ('DOCPROPERTY', 'MERGEFIELD'), '', '123.', '123.'),
    ('NoSwitch.testNone38', ('DOCPROPERTY', 'MERGEFIELD'), '', '-123.', '-123.'),
    ('NoSwitch.testNone39', ('DOCPROPERTY', 'MERGEFIELD'), '', '0', '0'),
    ('NoSwitch.testNone40', ('DOCPROPERTY', 'MERGEFIELD'), '', '00', '00'),
    ('NoSwitch.testNone41', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.', '0.'),
    ('NoSwitch.testNone42', ('DOCPROPERTY', 'MERGEFIELD'), '', '01', '01'),
    ('NoSwitch.testNone43', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.1', '0.1'),
    ('NoSwitch.testNone44', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.0', '0.0'),
    ('NoSwitch.testNone45', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.00', '0.00'),
    ('NoSwitch.testNone46', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.00 0', '0.00 0'),
    ('NoSwitch.testNone47', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.00 1', '0.00 1'),
    ('NoSwitch.testNone48', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.00 A', '0.00 A'),
    ('NoSwitch.testNone49', ('DOCPROPERTY', 'MERGEFIELD'), '', '0.00 W', '0.00 W'),
    ('NoSwitch.testNone50', ('DOCPROPERTY', 'MERGEFIELD'), '', '"01"', '"01"'),
    ('NoSwitch.testNone51', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.1"', '"0.1"'),
    ('NoSwitch.testNone52', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.0"', '"0.0"'),
    ('NoSwitch.testNone53', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.00"', '"0.00"'),
    ('NoSwitch.testNone54', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.00 0"', '"0.00 0"'),
    ('NoSwitch.testNone55', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.00 1"', '"0.00 1"'),
    ('NoSwitch.testNone56', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.00 A"', '"0.00 A"'),
    ('NoSwitch.testNone57', ('DOCPROPERTY', 'MERGEFIELD'), '', '"0.00 W"', '"0.00 W"'),
    ('NoSwitch.testNone58', ('DOCPROPERTY', 'MERGEFIELD'), '', '0000123456', '0000123456'),
    ('NoSwitch.testNone59', ('DOCPROPERTY', 'MERGEFIELD'), '', '000012345.006', '000012345.006'),
    ('NoSwitch.testNone60', ('DOCPROPERTY', 'MERGEFIELD'), '', '0000123AA456', '0000123AA456'),
    ('NoSwitch.testNone61', ('DOCPROPERTY', 'MERGEFIELD'), '', '0000123AA45.006', '0000123AA45.006'),
]  # fmt: skip


@pytest.mark.parametrize(
    ("field", "switch", "value", "expected"),
    [(f, switch, value, expected) for _test, fields, switch, value, expected in CASES for f in fields],
    ids=[f"{test}-{f}" for test, fields, *_rest in CASES for f in fields],
)
def test_docx4js_switch_cases(field, switch, value, expected):
    model = FldSimpleModel(f"{field} foo {switch}")
    if expected == REFUSED:
        with pytest.raises(FieldFormattingError):
            apply_formatting_switch(None, model, value)
    else:
        assert apply_formatting_switch(None, model, value) == expected


def test_the_case_table_is_docx4js():
    assert len(CASES) == 138


@pytest.mark.parametrize(
    ("switch", "value", "expected"),
    [
        ('\\@ "MMMM d, yyyy"', "2013-04-15", "April 15, 2013"),
        ('\\@ "MMMM d, yyyy HH:mm"', "2013-04-15T09:30:00", "April 15, 2013 09:30"),
        ("\\@ yyyy-MM-dd", "15/4/2013", "2013-04-15"),  # international: the day first
        ('\\@ "dddd, MMMM dd, yyyy"', "15/04/2013", "Monday, April 15, 2013"),
        ("\\@ M/d/yy", "15 April 2013", "4/15/13"),
        ('\\@ "h:mm am/pm"', "2013-04-15 13:05", "1:05 PM"),
        ("\\@ yyyy-MM-dd", "not a date", "not a date"),  # the switch has no effect
        ("\\@ yyyy", "2013-04-15T00:00:00Z", "2013-04-15T00:00:00Z"),  # docx4j infers no zone
        ('\\@ "d MMMM yyyy" \\* Upper', "15/04/2013", "15 APRIL 2013"),
    ],
)
def test_dates(switch, value, expected):
    assert apply_formatting_switch(None, FldSimpleModel(f"DOCPROPERTY foo {switch}"), value) == expected


def test_the_helpers():
    assert FldSimpleModel(' TITLE  \\* Upper ').name == "TITLE"
    assert FldSimpleModel(' QUOTE "a b" \\* Caps').parameters == ['"a b"', "\\*", "Caps"]
    assert FldSimpleModel("=2+3").name == "="
    # the spaces after a closing quote go inside it, as docx4j's convertDatePattern moves them
    assert convert_date_pattern("dddd d MMMM yyyy h:mm AM/PM 'at' HH") == "EEEE d MMMM yyyy K:mm a 'at 'HH"
    assert determine_date_format("2013-04-15t09:30") == "yyyy-MM-dd'T'HH:mm"
    assert determine_date_format("13/04/2013") == "dd/MM/yyyy"
    assert extract_number("€180,000.00 EUR") == "180000.00"
    # docx4j takes its grouping separator (a property, "," by default) out first, then the decimal symbol
    assert extract_number("1.234,5", ",") == "1.2345"
