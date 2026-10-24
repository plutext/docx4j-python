"""``effective(obj, name)``: the one reader of a value the file leaves out. CR-018 Phase C, section 4.4.

The model's fields never change: every one is ``None | T = None`` and stays None
when the attribute is absent (CR-001's ``SchemaDefaults=metadata``). ``effective``
reads the stated value, else the schema default the fork types once per field
(``XmlVar.schema_default_value``, stage 13 item 31), and only in the namespaces
whose defaults have been audited (Q10): ``w``, refreshed from docx4j ``46c727bf8``
(``tests/test_schema_defaults.py`` holds those defaults to docx4j's), ``w14`` and
``w15``. Anywhere else it raises ``LookupError``.
"""

from __future__ import annotations

import dataclasses
import inspect
import math
from importlib import import_module
from pathlib import Path

import pytest

from docx4j_py.model.values import AUDITED_NAMESPACES, effective, on_off

ROOT = Path(__file__).resolve().parents[2]


def _defaults(module: str) -> list[tuple[type, str, object]]:
    from docx4j_py.runtime import context

    mod = import_module(module)
    out = []
    for cls in vars(mod).values():
        if inspect.isclass(cls) and dataclasses.is_dataclass(cls) and cls.__module__ == module:
            for var in context().build(cls).get_attribute_vars():
                if var.schema_default is not None:
                    out.append((cls, var.name, var.schema_default_value))
    return out


def test_the_w_defaults_are_typed() -> None:
    """Every ``w`` default is a bool (52 fields), so no default in scope is a hazard ("", "NaN")."""
    defaults = _defaults("docx4j_py.wml")
    assert len(defaults) == 52
    assert {type(value) for _cls, _name, value in defaults} == {bool}
    for cls, name, value in defaults:
        assert effective(cls(), name) is value, (cls.__name__, name)


def test_the_w14_and_w15_defaults() -> None:
    """``w14``'s three are docx4j's patch of ST_OnOff, default true; ``w15``'s on/off elements take ``w:val``."""
    w14 = {(cls.__name__, name): value for cls, name, value in _defaults("docx4j_py.w14")}
    assert w14 == {("CTOnOff", "val"): True, ("ConflictMode", "val"): True, ("DiscardImageEditingData", "val"): True}
    from docx4j_py.runtime import context

    for cls, name, _value in _defaults("docx4j_py.w15"):
        var = next(v for v in context().build(cls).get_attribute_vars() if v.name == name)
        assert var.qname.startswith("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"), (cls, name)


def test_the_stated_value_wins_and_the_field_stays_none() -> None:
    from docx4j_py.wml import el

    fld = el.fldChar(fld_char_type="begin")
    assert fld.fld_lock is None
    assert effective(fld, "fld_lock") is False  # docx4j #697: absent is off
    assert effective(el.fldChar(fld_char_type="begin", fld_lock=True), "fld_lock") is True
    assert effective(el.fldChar(fld_char_type="begin", fld_lock=" false"), "fld_lock") == " false"  # kept, as read
    assert on_off(effective(el.fldChar(fld_char_type="begin", fld_lock=" false"), "fld_lock")) is True
    assert effective(el.b(), "val") is True
    assert effective(el.cols(), "equal_width") is True  # CR-059 kept it true
    assert effective(el.spacing(), "before_autospacing") is None  # CR-059: no default
    assert fld.fld_lock is None  # nothing was written


def test_outside_the_audited_namespaces_it_raises() -> None:
    """The hazards ("" and "NaN" defaults, 67 and 3 of them, all DrawingML or customXml) are out of scope."""
    from docx4j_py.dml.diagram import CTNumericRule
    from docx4j_py.runtime import context

    var = next(v for v in context().build(CTNumericRule).get_attribute_vars() if v.name == "val")
    assert math.isnan(var.schema_default_value)  # the fork types it; this package does not read it
    with pytest.raises(LookupError, match="not audited"):
        effective(CTNumericRule(), "val")
    with pytest.raises(LookupError, match="no attribute field"):
        effective(CTNumericRule(), "nonesuch")
    assert len(AUDITED_NAMESPACES) == 3


def test_a_bare_w14_checked_reads_checked() -> None:
    """CR-018 section 2.11 item 3: a bare ``<w14:checked/>`` reads checked, its schema default being true."""
    from docx4j_py.fragments import wml
    from docx4j_py.model.customxml.kinds import CheckboxContentControl

    w14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    for checked, reads in (
        ("<w14:checked/>", True),
        ('<w14:checked w14:val="0"/>', False),
        ('<w14:checked w14:val="1"/>', True),
        ('<w14:checked w14:val="off"/>', False),
        ("", False),  # no w14:checked at all: unchecked
    ):
        box = wml(f"<w14:checkbox {w14}>{checked}</w14:checkbox>")
        assert CheckboxContentControl(None, box).is_checked is reads, checked
