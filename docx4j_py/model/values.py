"""Values as Word 365 reads them: the one reader of each kind of value (CR-018).

The parse keeps a value as it was written; a reader here applies Word's rule to
it. That is the split CR-018 section 4.1a settles: the fork keeps what was
written, exactly and explicitly, and ``docx4j_py`` decides what it means. Under
``runtime.parser_config()``'s exact parse an on/off value that is not exactly
``true``, ``false``, ``1`` or ``0`` stays the string it was (``"off"``,
``" false"``, ``""``, ``"TRUE"``), so every reader of an on/off value must read
it here rather than ask whether it ``is not False``.

**ST_OnOff** (CR-018 section 1.1, rule 1, four rounds of probes in Word 365,
2026-10-24). A value is **off for exactly** ``false``, ``off`` or ``0``:
case-sensitive and untrimmed. It is **on for anything else**: absent (the
schema's true for an element's ``w:val``), ``true``, ``on``, ``1``, ``""``,
``" "``, ``TRUE``, ``FALSE``, ``OFF``, ``" false"``, ``"false "``, ``00``,
``no``, ``yes``. Word opens all of them without a repair prompt.

* :func:`on_off` reads one value, :func:`is_on` an on/off element such as
  ``w:b`` (absent is off, present is :func:`on_off` of its ``w:val``);
* :func:`xsd_boolean` reads ``xsd:boolean`` as XSD reads it, whitespace
  collapsed, for the namespaces whose booleans are not Word's ST_OnOff
  (DrawingML's ``c:`` and ``cx:``, ``adec:``).

Measured on ``b``, ``i``, ``caps``, ``smallCaps``, ``strike``, ``dstrike``,
``outline``, ``shadow``, ``emboss``, ``imprint`` and ``vanish``; every other
on/off value (``bCs``, ``w:keepNext``, ``w:tblHeader``, ``w:trackRevisions``,
the on/off attributes such as ``w:fldLock``, ``w14`` and ``w15``) is read by the
same rule by extension, as docx4j's ``OnOffAdapter`` reads all of them
(docx4j CR-060 phase 1, ``893658f68``).

:func:`on_off_normal_form` is what a save writes for a kept value (section 4.7,
:mod:`docx4j_py.openpackaging.normalise`).

CR-018 Phase B adds the enumerations' reader (``in_enumeration``) and the
colour reader (``colour``) here; Phase C adds ``effective()``, the one reader of
a value the file leaves out.
"""

from __future__ import annotations

from enum import Enum
from typing import Any

__all__ = ["OFF", "XSD_FALSE", "XSD_TRUE", "is_on", "on_off", "on_off_normal_form", "xsd_boolean"]

#: Rule 1: the only spellings Word reads as off. Exactly these, case-sensitive, untrimmed.
OFF: frozenset[str] = frozenset({"false", "off", "0"})

#: ``xsd:boolean``'s four lexical forms, after XSD's whitespace collapse.
XSD_TRUE: frozenset[str] = frozenset({"true", "1"})
XSD_FALSE: frozenset[str] = frozenset({"false", "0"})


def on_off(value: Any, default: bool = True) -> bool:
    """An ST_OnOff value as Word 365 reads it.

    ``None`` (the attribute absent) is `default`: the schema's true for an
    element's ``w:val`` (``BooleanDefaultTrue``), and the caller's answer for
    an attribute whose absence means something else (``w:fldLock``, absent, is
    off). A ``bool`` is itself, which is what the parse binds for exactly
    ``true``, ``false``, ``1`` and ``0``. An enumeration member reads by its
    value (``w:tblLook``'s ``on``/``off``, ``w14``'s and ``w15``'s ST_OnOff,
    OMML's ``m:val``). Any other value is off only if it is exactly
    ``"false"``, ``"off"`` or ``"0"``.
    """
    if value is None:
        return default
    if value is True or value is False:
        return value
    if isinstance(value, Enum):
        value = value.value
        if value is True or value is False:
            return value
    return str(value) not in OFF


def is_on(element: Any) -> bool:
    """An on/off element such as ``w:b`` or ``w:vanish``: absent is off; present, :func:`on_off` of its ``w:val``.

    A bare ``<w:b/>`` is on: ``w:val``'s schema default is true. A wildcard
    element (one the model holds untyped) reads its ``w:val`` attribute the
    same way.
    """
    if element is None:
        return False
    attributes = getattr(element, "attributes", None)
    if isinstance(attributes, dict) and not hasattr(element, "val"):
        value = next((v for k, v in attributes.items() if k == "val" or k.endswith("}val")), None)
        return on_off(value)
    return on_off(getattr(element, "val", None))


def on_off_normal_form(value: Any, absent_is_on: bool) -> bool | None:
    """What a save writes for an ST_OnOff value the parse kept as its text (CR-018 section 4.7's on/off rows).

    None means the attribute is dropped: Word reads the value as on and the
    attribute's absence is on too (an element's ``w:val``, whose schema default
    is true), so the value is written **bare**, as Word's own save writes it.
    Otherwise the ``bool`` Word reads, which the serialiser writes ``1`` or
    ``0``: off is ``0`` with the element kept, since an off value can override a
    style and a save looks nothing up.
    """
    on = on_off(value)
    if on and absent_is_on:
        return None
    return on


def xsd_boolean(value: Any, default: bool | None = None) -> bool | None:
    """An ``xsd:boolean`` value as XSD reads it: whitespace collapsed, then ``true``/``1`` or ``false``/``0``.

    For the namespaces whose booleans are XSD's, not Word's ST_OnOff:
    DrawingML's charts (``c:``, ``cx:``) and the decorative flag (``adec:``).
    ``None`` is `default`; a ``bool`` is itself; a value that is none of the
    four literals once collapsed is `default` too, as a schema-valid reader
    treats an invalid value as absent.
    """
    if value is None:
        return default
    if value is True or value is False:
        return value
    if isinstance(value, Enum):
        value = value.value
    text = " ".join(str(value).split())
    if text in XSD_TRUE:
        return True
    if text in XSD_FALSE:
        return False
    return default

