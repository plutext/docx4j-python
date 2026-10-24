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

:func:`effective` (CR-018 Phase C, section 4.4) is the one reader of a value the
file leaves out: the stated value, else the schema's default, typed. It answers
**absence**; the readers above answer **a present value's spelling**, so an
on/off attribute is ``on_off(effective(obj, "fld_lock"))``.

CR-018 Phase B adds the enumerations' reader (``in_enumeration``) and the
colour reader (``colour``) here.
"""

from __future__ import annotations

import threading
from enum import Enum
from typing import Any

__all__ = [
    "AUDITED_NAMESPACES",
    "OFF",
    "XSD_FALSE",
    "XSD_TRUE",
    "effective",
    "is_on",
    "on_off",
    "on_off_normal_form",
    "xsd_boolean",
]

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


# ---------------------------------------------------------------------------
# effective(): the value the file leaves out (CR-018 Phase C, section 4.4)
# ---------------------------------------------------------------------------

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"

#: The namespaces whose schema defaults have been audited, and only these (CR-018 Q10):
#: ``w``, by docx4j's #697 and CR-059 and the refresh from ``46c727bf8`` (52 fields, every
#: one ``true`` or ``false``); ``w14``, whose three are docx4j's own patch of ST_OnOff
#: (``w14:CT_OnOff``, ``conflictMode`` and ``discardImageEditingData``, default true, which
#: docx4j's getters answer for a bare element; unmeasured in Word, which always writes
#: ``w14:val``); and ``w15``, which declares none of its own (its on/off elements take
#: ``w:val``). Unaudited defaults are how docx4j's getters went wrong (#697, #641, the
#: autospacing bug, CR-059's ``anchorLock`` regression, the FontEmbedder bug), so anywhere
#: else :func:`effective` raises until someone audits that namespace.
AUDITED_NAMESPACES: frozenset[str] = frozenset({W, W14, W15})

_ATTRIBUTES: dict[type, dict[str, Any]] = {}
_ATTRIBUTES_LOCK = threading.Lock()


def _attribute_vars(cls: type) -> dict[str, Any]:
    """A class's attribute vars by field name, built once per class."""
    found = _ATTRIBUTES.get(cls)
    if found is None:
        from docx4j_py.runtime import context

        found = {var.name: var for var in context().build(cls).get_attribute_vars()}
        with _ATTRIBUTES_LOCK:
            _ATTRIBUTES[cls] = found
    return found


def effective(obj: Any, name: str) -> Any:
    """The value `obj` states for attribute `name`, else the schema's default, typed; else None.

    The stated value wins whatever it is, a kept ``str`` included (``" false"``,
    ``"red"``): the readers above decide what it means. Absent, it is the
    fork's ``XmlVar.schema_default_value`` (stage 13, item 31), the schema
    default converted once as a parse would convert it: ``True`` for
    ``"true"``, an enumeration's member. The field itself never changes: it
    stays None when absent, and nothing but this reads the default.

    Raises:
        LookupError: `name` is not an attribute of `obj`'s class, or the
            attribute's namespace is not in :data:`AUDITED_NAMESPACES`
            (CR-018 Q10).
    """
    var = _attribute_vars(type(obj)).get(name)
    if var is None:
        raise LookupError(f"{type(obj).__name__} has no attribute field {name!r}")
    qname = var.qname or ""
    namespace = qname[1:].split("}", 1)[0] if qname.startswith("{") else ""
    if namespace not in AUDITED_NAMESPACES:
        raise LookupError(
            f"{type(obj).__name__}.{name}: the schema defaults of {namespace or 'no namespace'!r} are not audited "
            "(CR-018 Q10); read the field, which is None when the attribute is absent"
        )
    value = getattr(obj, name, None)
    if value is not None:
        return value
    return var.schema_default_value
