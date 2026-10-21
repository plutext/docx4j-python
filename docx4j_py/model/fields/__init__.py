"""Fields: docx4j's ``org.docx4j.model.fields``, the parts the text model computes with. CR-011 section 6.

* :class:`FldSimpleModel`: a field instruction's name and parameters.
* :func:`apply_formatting_switch`: docx4j's ``FormattingSwitchHelper``, the
  ``\\@`` (date and time), ``\\#`` (number) and ``\\*`` (capitalisation,
  ``Arabic``, ``Roman``) switches on a field's value --- a measured subset,
  not a general engine (CR-011 risk 5).
* :class:`DocPropertyResolver` and :class:`DocVariableResolver`: a
  ``DOCPROPERTY``'s and a ``DOCVARIABLE``'s value, read without changing the
  package.

The text model's ``fields: update`` (WD REQ-506) computes through these.
Imported on first use: ``import docx4j_py`` does not import it.
"""

from __future__ import annotations

from docx4j_py.model.fields.formatting import (
    FieldFormattingError,
    FieldValueError,
    FldSimpleModel,
    apply_formatting_switch,
    convert_date_pattern,
    determine_date_format,
    extract_number,
    find_first_switch_value,
)
from docx4j_py.model.fields.resolvers import DocPropertyResolver, DocVariableResolver

__all__ = [
    "DocPropertyResolver",
    "DocVariableResolver",
    "FieldFormattingError",
    "FieldValueError",
    "FldSimpleModel",
    "apply_formatting_switch",
    "convert_date_pattern",
    "determine_date_format",
    "extract_number",
    "find_first_switch_value",
]
