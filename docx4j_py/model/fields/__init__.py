"""Fields: docx4j's ``org.docx4j.model.fields``, the parts the text model computes with. CR-011 section 6.

* :class:`FieldInstruction`: a field instruction's name, parameters,
  argument and switches (docx4j's, CR-067 D2; CR-025 Phase B), and
  :class:`FldSimpleModel`, a subclass that keeps docx4j's older name.
* :func:`apply_formatting_switch`: docx4j's ``FormattingSwitchHelper``, the
  ``\\@`` (date and time), ``\\#`` (number) and ``\\*`` (capitalisation,
  ``Arabic``, ``Roman``) switches on a field's value --- a measured subset,
  not a general engine (CR-011 risk 5).
* :class:`DocPropertyResolver` and :class:`DocVariableResolver`: a
  ``DOCPROPERTY``'s and a ``DOCVARIABLE``'s value, read without changing the
  package.
* :func:`evaluate` and :func:`register`, over
  :mod:`~docx4j_py.model.fields.evaluators`: docx4j's ``FieldEvaluators`` at
  ``c3e9920e5`` (its CR-067 phase 1; CR-025 Phase B), one evaluator per field
  name, with :class:`FieldContext` (the package, the zone, a clock, the user,
  the section, the pass's counters) and :class:`FieldUser`.

The text model's ``fields: update`` (WD REQ-506) computes through these.
Imported on first use: ``import docx4j_py`` does not import it.
"""

from __future__ import annotations

from docx4j_py.model.fields.context import FieldContext, FieldUser
from docx4j_py.model.fields.evaluators import evaluate, register
from docx4j_py.model.fields.formatting import (
    FieldFormattingError,
    FieldValueError,
    FldSimpleModel,
    apply_formatting_switch,
    convert_date_pattern,
    determine_date_format,
    extract_number,
    find_first_switch_value,
    format_date,
)
from docx4j_py.model.fields.instruction import (
    FieldInstruction,
    split_parameters,
    unquote,
)
from docx4j_py.model.fields.resolvers import DocPropertyResolver, DocVariableResolver
from docx4j_py.model.fields.sources import parse_w3cdtf

__all__ = [
    "DocPropertyResolver",
    "DocVariableResolver",
    "FieldContext",
    "FieldFormattingError",
    "FieldInstruction",
    "FieldUser",
    "FieldValueError",
    "FldSimpleModel",
    "apply_formatting_switch",
    "convert_date_pattern",
    "determine_date_format",
    "evaluate",
    "extract_number",
    "find_first_switch_value",
    "format_date",
    "parse_w3cdtf",
    "register",
    "split_parameters",
    "unquote",
]
