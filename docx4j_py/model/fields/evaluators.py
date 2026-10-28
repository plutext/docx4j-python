"""The field evaluators by name, and the one way to evaluate a field through them. CR-025 Phase B, D5.

docx4j's ``org.docx4j.model.fields.FieldEvaluators`` at ``c3e9920e5`` (its CR-067
D1, phase 1): an evaluator's result with the formatting switches applied, or
None where the field keeps its cached result (docx4j's D5). A caller may
:func:`register` one of its own, for a field this does not compute or to
replace one it does::

    from docx4j_py.model.fields import FieldContext, FieldUser, evaluate

    context = FieldContext(pkg, user=FieldUser("Jo", "J"))
    evaluate(" USERNAME \\\\* Upper ", context)       # 'JO'
    evaluate(" CREATEDATE ", context)                # '2/3/2024 3:05:00 PM', in the package's zone

Registered here, as docx4j registers them:

* ``DOCPROPERTY`` and ``DOCVARIABLE``;
* the document-information fields ``AUTHOR``, ``COMMENTS``, ``KEYWORDS``,
  ``LASTSAVEDBY``, ``SUBJECT``, ``TITLE`` and ``REVNUM`` (one given a text
  argument shows that text, as Word sets the property from it);
* ``CREATEDATE`` and ``SAVEDATE``: the core property's instant, to the
  minute, in the context's zone; with ``\\@`` its picture, without one Word's
  default, the short date and the long time (``2/3/2024 3:05:00 PM``,
  measured in Word 365, docx4j CR-067 section 11.3); a date alone or a time
  with no offset UTC, as Word reads them (CR-025 D8, D9 as Phase B2 made it;
  docx4j's D13);
* ``SEQ`` (counted in the context's pass; ``\\r``, ``\\c``, ``\\h``; not
  ``\\s``); ``REF`` to a bookmark the context has read (with ``\\n``,
  ``\\r``, ``\\w``, ``\\p`` or ``\\f``, which need numbering or layout, the
  cached result); ``QUOTE`` (no argument is empty, measured);
* ``DATE``, ``TIME`` and ``PRINTDATE`` at the context's clock in its zone;
* ``INFO <name>`` as the field it names, where the context's consumer
  evaluates that field (CR-025 D12); ``SECTION`` where the context says which;
  ``NOTEREF`` to a bookmark on a note reference; ``USERNAME``,
  ``USERINITIALS`` and ``USERADDRESS`` from the context's user.

Not here: ``MERGEFIELD``, ``FORMTEXT``, the layout fields (``PAGE``,
``NUMPAGES``, ``PAGEREF``, ``STYLEREF``, ``TOC``) and the external ones.
"""

from __future__ import annotations

import threading
from collections.abc import Callable
from typing import Any

from docx4j_py.model.fields.formatting import (
    FieldFormattingError,
    FieldValueError,
    apply_formatting_switch,
    format_date,
)
from docx4j_py.model.fields.instruction import FieldInstruction

__all__ = ["DOC_INFO", "Evaluator", "evaluate", "formats_itself", "get", "has", "names", "register"]

#: An evaluator: the field's result, or None to keep its cached one.
Evaluator = Callable[[FieldInstruction, Any], "str | None"]

#: The document-information fields and the property each shows, as ``DOCPROPERTY`` names it (docx4j ``DOC_INFO``).
DOC_INFO: dict[str, str] = {
    "AUTHOR": "Author",
    "COMMENTS": "Comments",
    "KEYWORDS": "Keywords",
    "LASTSAVEDBY": "LastSavedBy",
    "SUBJECT": "Subject",
    "TITLE": "Title",
    "REVNUM": "RevisionNumber",
    "CREATEDATE": "CreateTime",
    "SAVEDATE": "LastSavedTime",
}

_REF_UNSUPPORTED = frozenset({"\\n", "\\r", "\\w", "\\p", "\\f"})

_REGISTRY: dict[str, tuple[Evaluator, bool]] = {}
_LOCK = threading.Lock()


def register(name: str, evaluator: Evaluator, *, formats_itself: bool = False) -> None:
    """Register (or replace) the evaluator for a field name, in any case.

    `formats_itself`: the evaluator reads its own switches (the date fields'
    ``\\@``), so :func:`evaluate` applies none after it.
    """
    if not name or evaluator is None:
        raise ValueError("a name and an evaluator")
    with _LOCK:
        _REGISTRY[name.upper()] = (evaluator, formats_itself)


def get(name: str | None) -> Evaluator | None:
    """The evaluator for the field name, or None."""
    entry = _REGISTRY.get(name.upper()) if name else None
    return entry[0] if entry is not None else None


def formats_itself(name: str | None) -> bool:
    """Whether the field name's evaluator reads its own switches."""
    entry = _REGISTRY.get(name.upper()) if name else None
    return entry is not None and entry[1]


def has(name: str | None) -> bool:
    """Whether a field of that name is evaluated here."""
    return get(name) is not None


def names() -> frozenset[str]:
    """The names evaluated here, upper case."""
    return frozenset(_REGISTRY)


def evaluate(instruction: FieldInstruction | str | None, context: Any) -> str | None:
    """The field's result as it is to be shown, or None where it keeps its cached result.

    None for no evaluator of its name, or an evaluator that could not compute
    it (a property that does not exist, a bookmark not read). The switches
    are applied after the evaluator, unless it formats itself.

    Raises:
        FieldFormattingError: a picture docx4j refuses (``\\# x#.#x``), as
            docx4j's ``FormattingSwitchHelper`` throws it.
    """
    if instruction is None or context is None:
        return None
    if not isinstance(instruction, FieldInstruction):
        instruction = FieldInstruction(instruction)
    entry = _REGISTRY.get(instruction.name or "")
    if entry is None:
        return None
    evaluator, own_switches = entry
    try:
        value = evaluator(instruction, context)
    except FieldFormattingError:
        raise
    except FieldValueError:
        return None
    except Exception:  # noqa: BLE001 - a field whose value cannot be found keeps its cached result
        return None
    if value is None or own_switches:
        return value
    return apply_formatting_switch(context.package, instruction, value, context.language)


# ---------------------------------------------------------------------------
# the evaluators
# ---------------------------------------------------------------------------


def _doc_property(instruction: FieldInstruction, context: Any) -> str | None:
    key = instruction.argument
    return None if key is None else context.properties().value(key)


def _doc_variable(instruction: FieldInstruction, context: Any) -> str | None:
    key = instruction.argument
    return None if key is None else context.variables().value(key)


def _doc_info(property_name: str) -> Evaluator:
    def evaluator(instruction: FieldInstruction, context: Any) -> str | None:
        argument = instruction.argument
        if argument is not None:
            return argument  # AUTHOR "Jo": Word sets the property from it, and shows it
        return context.properties().value(property_name)

    return evaluator


def _sequence(instruction: FieldInstruction, context: Any) -> str | None:
    identifier = instruction.argument
    if identifier is None:
        return None
    key = identifier.upper()
    n = context.sequences.get(key, 0)
    parameters = instruction.parameters
    reset = instruction.switch_value("\\r") if instruction.has_switch("\\r") else None
    if reset:
        try:
            n = int(reset.strip())
        except ValueError:
            return None
    elif "\\c" not in parameters:
        n += 1
    context.sequences[key] = n
    if "\\h" in parameters:
        return ""
    return str(n)


def _ref(instruction: FieldInstruction, context: Any) -> str | None:
    bookmark = instruction.argument
    if bookmark is None:
        return None
    if any(p.lower() in _REF_UNSUPPORTED for p in instruction.parameters):
        return None
    return context.bookmarks.get(bookmark.lower())


def _quote(instruction: FieldInstruction, context: Any) -> str | None:
    # no argument: empty, as Word 365 shows it (docx4j CR-067 section 11.3, case 12)
    argument = instruction.argument
    return "" if argument is None else argument


def _clock_date(instruction: FieldInstruction, context: Any) -> str | None:
    from docx4j_py.model.dates import wall_clock

    return format_date(instruction, wall_clock(context.now(), context.zone))


#: Word's default for ``CREATEDATE`` and ``SAVEDATE`` with no ``\\@``: the short date and the long time, US English.
WORD_DEFAULT_DATE_TIME = "M/d/yyyy h:mm:ss a"


def _core_date(property_name: str) -> Evaluator:
    """``CREATEDATE`` and ``SAVEDATE``: docx4j's ``CoreDate`` (CR-025 D8).

    The core property is an instant (:func:`~docx4j_py.model.fields.sources.parse_w3cdtf`:
    a value with no offset, or a date alone, is UTC, as Word reads it; docx4j's
    D13, CR-025 Phase B2), cut to the minute and shown in the context's zone.
    """

    def evaluator(instruction: FieldInstruction, context: Any) -> str | None:
        from docx4j_py.model.dates import to_the_minute, wall_clock
        from docx4j_py.model.fields.sources import parse_w3cdtf
        from docx4j_py.model.java_text import simple_date_format

        when = parse_w3cdtf(context.properties().value(property_name))
        if when is None:
            return None
        # to the minute, as Word reads a file's time (04:05:06Z shows as :05:00), then the zone (D12)
        clock = wall_clock(to_the_minute(when), context.zone)
        if instruction.has_switch("\\@"):
            return format_date(instruction, clock)
        return simple_date_format(WORD_DEFAULT_DATE_TIME, clock)

    return evaluator


def _info(instruction: FieldInstruction, context: Any) -> str | None:
    # INFO <name> [switches]: the named document-information field itself
    text = instruction.parameter_text
    if instruction.argument is None or text is None:
        return None
    named = FieldInstruction(text)
    if named.name is None or named.name == "INFO" or not has(named.name):
        return None
    evaluates = getattr(context, "evaluates", None)
    if evaluates is not None and not evaluates(named.name):
        return None  # only a field the consumer computes (CR-025 D12; docx4j asks the whole registry)
    return evaluate(named, context)


def _section(instruction: FieldInstruction, context: Any) -> str | None:
    section = getattr(context, "section", None)
    return None if section is None else str(section)


def _note_ref(instruction: FieldInstruction, context: Any) -> str | None:
    bookmark = instruction.argument
    return None if bookmark is None else context.note_numbers().get(bookmark.lower())


def _user_name(instruction: FieldInstruction, context: Any) -> str | None:
    user = context.user
    return None if user is None else user.name


def _user_initials(instruction: FieldInstruction, context: Any) -> str | None:
    # a user supplied without initials or an address shows them empty, as Word shows a blank one
    user = context.user
    return None if user is None else (user.initials or "")


def _user_address(instruction: FieldInstruction, context: Any) -> str | None:
    user = context.user
    return None if user is None else (user.address or "")


register("DOCPROPERTY", _doc_property)
register("DOCVARIABLE", _doc_variable)
for _name, _property in DOC_INFO.items():
    register(_name, _doc_info(_property))
register("SEQ", _sequence)
register("REF", _ref)
register("QUOTE", _quote)
register("DATE", _clock_date, formats_itself=True)
register("TIME", _clock_date, formats_itself=True)
register("PRINTDATE", _clock_date, formats_itself=True)  # the conversion is the printing, as docx4j's
# tier 1 (docx4j CR-067 phase 1): CREATEDATE and SAVEDATE replace the document-information rule
register("CREATEDATE", _core_date("CreateTime"), formats_itself=True)
register("SAVEDATE", _core_date("LastSavedTime"), formats_itself=True)
register("INFO", _info)
register("SECTION", _section)
register("NOTEREF", _note_ref)
register("USERNAME", _user_name)
register("USERINITIALS", _user_initials)
register("USERADDRESS", _user_address)
del _name, _property
