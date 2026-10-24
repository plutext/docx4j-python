"""The ``w`` namespace's schema defaults, as the model carries them and as docx4j's ``wml.xsd`` declares them.

CR-018 section 4.5 item 4. The copy in ``schemas/wml/wml.xsd`` is refreshed from
docx4j ``46c727bf8`` (#697 ``77ba784d6``, CR-059 ``ac97b56c5`` and ``46c727bf8``):
the 2008 on/off defaults corrected, 36 field defaults ``true`` to ``false`` and
10 gone. Nothing reads a default until CR-018 Phase C's ``effective()``, so the
refresh changes no answer; these tests hold the defaults themselves, so that the
next refresh is checked mechanically:

* every ``w`` attribute default in the model is the one the copy declares, for
  the attribute of the schema type the class was generated from (the type's own
  attributes, its attribute groups' and its extension base's), and every
  attribute name's defaults over the whole schema are the model's;
* the copy declares the defaults docx4j's ``wml.xsd`` does at ``46c727bf8``,
  read from the sibling checkout's history (skipped when it is not there).
"""

from __future__ import annotations

import dataclasses
import inspect
import subprocess
from pathlib import Path

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
COPY = ROOT / "schemas" / "wml" / "wml.xsd"
DOCX4J = ROOT.parent / "docx4j"
#: The docx4j commit the copy of ``wml.xsd`` is from (``schemas/PATCHES.md``).
REFRESHED_AT = "46c727bf8"

XS = "{http://www.w3.org/2001/XMLSchema}"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _local(qname: str | None) -> str | None:
    return None if qname is None else qname.split(":")[-1]


class _Schema:
    """``wml.xsd``'s attribute declarations, by container, with what each named type resolves to."""

    def __init__(self, data: bytes) -> None:
        self.root = etree.fromstring(data)
        self.types = {e.get("name"): e for e in self.root.findall(XS + "complexType")}
        self.groups = {e.get("name"): e for e in self.root.findall(XS + "attributeGroup")}

    def own(self, node: etree._Element) -> dict[str, str | None]:
        """The attributes `node` declares itself or through attribute groups, name to default."""
        out: dict[str, str | None] = {}
        for child in node:
            if not isinstance(child.tag, str):
                continue
            if child.tag == XS + "attribute" and child.get("name") is not None:
                out[child.get("name")] = child.get("default")
            elif child.tag == XS + "attributeGroup" and child.get("ref") is not None:
                group = self.groups.get(_local(child.get("ref")))
                if group is not None:
                    out.update(self.own(group))
            elif child.tag in (XS + "complexContent", XS + "simpleContent"):
                for derivation in child:
                    if isinstance(derivation.tag, str):
                        out.update(self.own(derivation))
        return out

    def resolved(self, type_name: str, seen: frozenset[str] = frozenset()) -> dict[str, str | None] | None:
        """A named complex type's attributes, its extension base's included; None if it is not declared."""
        node = self.types.get(type_name)
        if node is None or type_name in seen:
            return None
        out: dict[str, str | None] = {}
        for derivation in node.iter(XS + "extension"):
            base = self.resolved(_local(derivation.get("base")) or "", seen | {type_name})
            if base is not None:
                out.update(base)
        out.update(self.own(node))
        return out

    def anonymous(self, element_name: str) -> list[dict[str, str | None]]:
        """The attributes of each element of that name declared with a type of its own, its base's included."""
        out = []
        for node in self.root.iter(XS + "element"):
            inline = node.find(XS + "complexType")
            if node.get("name") != element_name or inline is None:
                continue
            found: dict[str, str | None] = {}
            for derivation in inline.iter(XS + "extension"):
                found.update(self.resolved(_local(derivation.get("base")) or "") or {})
            found.update(self.own(inline))
            out.append(found)
        return out

    def every(self) -> dict[tuple[str, str], str | None]:
        """Every attribute declaration, keyed by where it is (a named type or group, or an element path) and its name."""
        out: dict[tuple[str, str], str | None] = {}
        for node in self.root.iter(XS + "attribute"):
            name = node.get("name")
            if name is None:
                continue
            where = []
            for ancestor in node.iterancestors():
                if ancestor.get("name") is not None:
                    where.append(f"{etree.QName(ancestor).localname}:{ancestor.get('name')}")
            out[("/".join(reversed(where)), name)] = node.get("default")
        return out

    def by_name(self) -> dict[str, set[str | None]]:
        """Each attribute name's defaults over the whole schema (None for a declaration with none)."""
        out: dict[str, set[str | None]] = {}
        for (_where, name), default in self.every().items():
            out.setdefault(name, set()).add(default)
        return out


def _classes() -> list[type]:
    from docx4j_py import wml

    return [
        cls
        for cls in vars(wml).values()
        if inspect.isclass(cls) and dataclasses.is_dataclass(cls) and cls.__module__ == "docx4j_py.wml"
    ]


def _element_names() -> dict[type, set[str]]:
    """The element names each class is bound to by a field of another (an anonymous type's class has no type name)."""
    from docx4j_py import runtime

    out: dict[type, set[str]] = {}
    for cls in _classes():
        for var in runtime.context().build(cls).get_element_vars():
            for each in (var, *(var.elements or {}).values()):
                for bound in each.types:
                    if dataclasses.is_dataclass(bound):
                        out.setdefault(bound, set()).add(each.local_name)
    return out


def _model() -> list[tuple[type, str, str, str | None]]:
    """Every ``w`` attribute of every class of ``docx4j_py.wml``: (class, schema type name, attribute name, default)."""
    from docx4j_py import runtime

    out = []
    for cls in _classes():
        meta = runtime.context().build(cls)
        type_name = etree.QName(meta.qname).localname if meta.qname else cls.__name__
        for var in meta.get_attribute_vars():
            if var.namespaces and W not in var.namespaces:
                continue
            if not var.namespaces and not (var.qname or "").startswith("{" + W + "}"):
                continue
            out.append((cls, type_name, var.local_name, var.schema_default))
    return out


@pytest.fixture(scope="module")
def copy() -> _Schema:
    return _Schema(COPY.read_bytes())


def test_each_model_default_is_the_copy_s_for_its_type(copy: _Schema) -> None:
    """A class carries its schema type's defaults, attribute by attribute.

    The type is the named complex type the class was generated from, or, for a
    class of an anonymous type, every element of that name declared with a type
    of its own (all of which must agree).
    """
    elements = _element_names()
    checked = stated = 0
    wrong = []
    for cls, type_name, attribute, default in _model():
        stated += default is not None
        declared = copy.resolved(type_name)
        candidates = [declared] if declared is not None and attribute in declared else []
        if not candidates:
            for name in elements.get(cls, ()):
                candidates.extend(found for found in copy.anonymous(name) if attribute in found)
        if not candidates:
            continue
        checked += default is not None
        values = {found[attribute] for found in candidates}
        if values != {default}:
            wrong.append(f"{cls.__name__}.{attribute}: model {default!r}, wml.xsd {sorted(map(str, values))}")
    assert not wrong, "\n".join(wrong)
    assert stated == 52 and checked >= 41, (stated, checked)  # 41 of the 52 by type; every one by name, below


def test_each_attribute_name_has_the_copy_s_defaults(copy: _Schema) -> None:
    """Over the whole schema, each attribute name's set of defaults is the model's (anonymous types included)."""
    model: dict[str, set[str | None]] = {}
    for _cls, _type, attribute, default in _model():
        model.setdefault(attribute, set()).add(default)
    schema = copy.by_name()
    wrong = [
        f"{name}: model {sorted(map(str, model.get(name, set())))}, wml.xsd {sorted(map(str, values))}"
        for name, values in sorted(schema.items())
        if any(v is not None for v in values) and model.get(name, set()) - {None} != values - {None}
    ]
    assert not wrong, "\n".join(wrong)


def test_the_refreshed_defaults() -> None:
    """The refresh's own numbers: 52 ``w`` defaults on the model's fields (62 before), none of #697's left true."""
    model = [(cls.__name__, attribute, default) for cls, _t, attribute, default in _model()]
    stated = [row for row in model if row[2] is not None]
    by_field = {(name, attribute): default for name, attribute, default in model}
    assert by_field[("FldChar", "fldLock")] == "false"  # docx4j #697, 77ba784d6
    assert by_field[("FldChar", "dirty")] == "false"
    assert by_field[("CTDocProtect", "enforcement")] == "false"
    assert by_field[("CTFramePr", "anchorLock")] == "false"  # CR-059
    assert by_field[("Style", "customStyle")] == "false"
    assert by_field[("CTColumns", "equalWidth")] == "true"
    assert by_field[("CtPprBaseSpacing", "beforeAutospacing")] is None  # CR-059: no default
    assert by_field[("CTTrackChangesView", "insDel")] is None
    assert by_field[("BooleanDefaultTrue", "val")] == "true"
    assert by_field[("BooleanDefaultFalse", "val")] == "false"
    assert len(stated) == 52, len(stated)


def _docx4j_wml() -> bytes:
    if not (DOCX4J / ".git").exists():
        pytest.skip("../docx4j is not checked out")
    try:
        return subprocess.run(
            ["git", "-C", str(DOCX4J), "show", f"{REFRESHED_AT}:xsd/wml/wml.xsd"],
            check=True,
            capture_output=True,
        ).stdout
    except subprocess.CalledProcessError:
        pytest.skip(f"../docx4j has no commit {REFRESHED_AT}")


def test_the_copy_declares_docx4j_s_defaults_at_the_refresh(copy: _Schema) -> None:
    """Every attribute declaration of docx4j's ``wml.xsd`` at ``46c727bf8`` has the copy's default, and no other."""
    theirs = _Schema(_docx4j_wml()).every()
    ours = copy.every()
    # the copy's patches (schemas/PATCHES.md) promote CT_Dir and CT_Bdo (patch 1) and name CT_StylePaneFilter
    # (patch 3): keyed by where they are, those attributes sit elsewhere, and none of them has a default
    moved = {key for key in ours if key not in theirs}
    patched = ("complexType:CT_Dir", "complexType:CT_Bdo", "complexType:CT_StylePaneFilter")  # patches 1 and 3
    assert all(key[0].startswith(patched) for key in moved), sorted(moved)
    differ = [
        f"{where} @{name}: copy {ours[(where, name)]!r}, docx4j {default!r}"
        for (where, name), default in sorted(theirs.items())
        if (where, name) in ours and ours[(where, name)] != default
    ]
    assert not differ, "\n".join(differ)
    gone = sorted(key for key in theirs if key not in ours)
    assert all(where.endswith(("element:dir", "element:bdo")) for where, _ in gone), gone
    assert all(theirs[key] is None for key in gone) and all(ours[key] is None for key in moved), (gone, moved)
