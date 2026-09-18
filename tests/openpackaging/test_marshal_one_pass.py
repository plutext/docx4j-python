"""One pass instead of three: the acceptance test of the fork's CR-001 Phase A.

``XmlPart._marshal`` used to render the object to a **string**, parse that
string back into an lxml tree, clean the namespaces, copy the whole tree into a
new root to add the prefixes ``mc:Ignorable`` names, and serialise again. On a
771 KB main part that copy alone was 345 ms of an 820 ms marshal. It now renders
straight into a tree and serialises once.

The acceptance of that change is **no change in the bytes**, so
:func:`reference_marshal` below is the old method, copied from
``xml_part.py`` at commit 959edcb, and every typed XML part of every sample is
marshalled both ways and compared byte for byte.

One deliberate departure, and it is why the reference has a ``sorted()`` the
original did not: the original iterated a ``set`` of prefixes, so the *order*
of the namespace declarations it re-added depended on ``PYTHONHASHSEED`` --- 49
of the 150 typed parts here had no single answer between runs. The bytes are
the same bytes in a different order (canonically identical either way, which is
why the round trip never saw it); the new method sorts, and this test's
reference sorts with it, so the comparison is about the one-pass change and
nothing else.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_marshal_one_pass.py -q
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py.namespaces import PREFIXES  # noqa: E402
from docx4j_py.openpackaging import OpcPackage, XmlPart  # noqa: E402
from docx4j_py.openpackaging.parts.xml_part import (  # noqa: E402
    XML_DECLARATION,
    _element_name,
    _ignorable_prefixes,
)

SAMPLES = sorted(p for p in (ROOT / "samples").iterdir() if p.suffix.lower() != ".md")

_NS_MAP: dict[str, str] = {p: u for p, u in PREFIXES.items() if p != "xml"}

_LXML_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)

logger = logging.getLogger("docx4j_py.openpackaging")


# ---------------------------------------------------------------------------
# the reference implementation: XmlPart._marshal as it was before Phase A
# ---------------------------------------------------------------------------


def _rename(root: etree._Element, name: str) -> etree._Element:
    """A copy of `root` under a different element name, children moved over."""
    new = etree.Element(name, nsmap=root.nsmap)
    for key, value in root.attrib.items():
        new.set(key, value)
    new.text = root.text
    for child in list(root):
        new.append(child)
    return new


def _redeclare(root: etree._Element, ns_map: dict[str, str]) -> etree._Element:
    """A copy of `root` whose namespace declarations are exactly `ns_map`."""
    new = etree.Element(root.tag, nsmap=ns_map)
    for key, value in root.attrib.items():
        new.set(key, value)
    new.text = root.text
    for child in list(root):
        new.append(child)
    return new


def reference_marshal(part: XmlPart, obj: Any) -> bytes:
    """``XmlPart._marshal`` as it stood at 959edcb: render, re-parse, copy, write."""
    from docx4j_py.runtime import serializer

    ns_map = _NS_MAP
    if part.default_namespace is not None:
        ns_map = {p: u for p, u in _NS_MAP.items() if u != part.default_namespace}
        ns_map[None] = part.default_namespace  # type: ignore[index]
    text = serializer(bool_format=part.bool_format).render(obj, ns_map)
    root = etree.fromstring(text.encode("utf-8"), _LXML_PARSER)

    wanted = part.root_name or _element_name(obj)
    if wanted is not None and root.tag != wanted:
        root = _rename(root, wanted)

    etree.cleanup_namespaces(root)
    ns_map = dict(root.nsmap)
    for prefix in sorted(_ignorable_prefixes(root)):  # sorted: see the module docstring
        if prefix in ns_map:
            continue
        uri = PREFIXES.get(prefix) or part.source_ns_map.get(prefix)
        if uri is not None:
            ns_map[prefix] = uri
    if ns_map != root.nsmap:
        root = _redeclare(root, ns_map)

    return XML_DECLARATION + etree.tostring(root, encoding="utf-8")


# ---------------------------------------------------------------------------
# the acceptance: every typed part of every sample, both ways
# ---------------------------------------------------------------------------


def typed_parts(package: OpcPackage) -> list[XmlPart]:
    return [
        part
        for part in package.parts.parts()
        if isinstance(part, XmlPart) and part.model_class is not None
    ]


@pytest.mark.parametrize("docx", SAMPLES, ids=lambda p: p.name)
def test_one_pass_marshal_is_byte_for_byte_the_old_one(docx, caplog):
    """Every typed part of every sample marshals to exactly the old bytes."""
    checked = 0
    with caplog.at_level(logging.ERROR, logger="docx4j_py.openpackaging"):
        with OpcPackage.load(docx) as package:
            for part in typed_parts(package):
                contents = part.contents
                assert part._marshal(contents) == reference_marshal(part, contents), (
                    f"{docx.name} {part.part_name}"
                )
                checked += 1
    assert checked > 0


def test_the_namespace_declarations_no_longer_depend_on_the_hash_seed():
    """The re-added ``mc:Ignorable`` prefixes are declared in a fixed order.

    ``Normal.dotm``'s main part names four prefixes, of which more than one is
    declared by nothing in the tree, so the old ``set`` iteration made the root
    element's bytes a function of ``PYTHONHASHSEED``.
    """
    with OpcPackage.load(ROOT / "samples" / "Normal.dotm") as package:
        part = package.main_document_part
        first = part._marshal(part.contents)
    root = etree.fromstring(first.split(b"\r\n", 1)[1])
    declared = [p for p in root.nsmap if p is not None]
    ignorable = sorted(_ignorable_prefixes(root))
    assert ignorable == ["w14", "w15", "w16se", "wp14"]
    # the ones nothing in the tree uses are declared last, in sorted order
    tail = [p for p in declared if p in ignorable]
    assert tail == sorted(tail)


def test_a_renamed_root_keeps_its_namespace_declarations():
    """``root_name`` differing from the object's element name is an in-place rename.

    The glossary document part is a ``Document`` under ``w:glossaryDocument``,
    which is the only shape in the corpus where the two differ.
    """
    with OpcPackage.load(ROOT / "samples" / "Normal.dotm") as package:
        part = package.parts["/word/glossary/document.xml"]
        data = part._marshal(part.contents)
    root = etree.fromstring(data.split(b"\r\n", 1)[1])
    assert etree.QName(root).localname == "glossaryDocument"
    assert root.nsmap["w"] == PREFIXES["w"]


def test_an_ignorable_prefix_nobody_can_resolve_is_warned_about(caplog):
    """The warning of CR-002 section 5.6 survives the rewrite."""
    from docx4j_py.openpackaging import MainDocumentPart
    from docx4j_py.wml import el
    from docx4j_py.wml.builders import p as make_p

    part = MainDocumentPart("/word/document.xml")
    document = el.document(body=el.body(content=[make_p("hi")]))
    document.ignorable = "nosuchprefix"
    part.set_contents(document)
    with caplog.at_level(logging.WARNING, logger="docx4j_py.openpackaging"):
        data = part.xml
    assert b"nosuchprefix" in data
    assert any("nosuchprefix" in record.getMessage() for record in caplog.records)
