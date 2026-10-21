"""Where a save declares the prefixes ``mc`` attributes name. CR-011 Q3 and Q11, docx4j CR-052.

The port of docx4j-core-tests' ``org.docx4j.jaxb.McSourcePrefixesTest`` at
``5df2e253e`` (CR-052 phase 1), less the cases about Java's marshal paths (the
stream writer, a DOM, Flat OPC, the binder, the canonicalising path): every
prefix ``Requires``, ``mc:Ignorable`` or ``mc:MustUnderstand`` names in a
saved ``word/document.xml`` resolves where it is named, and is declared where
Word 365 puts it --- a prefix the table binds as the source did on the root,
any other on the ``mc:Choice`` naming it, ahead of ``Requires``, a prefix the
root's ``mc:Ignorable`` names once on the root, nothing on ``mc:Fallback``.

Before CR-011 a re-marshalled save left the probe's ``zz9`` (three Choices)
and ``shape2010`` (one) undeclared, and ``waterfall.docx``'s ``cx1``, which
Word 365 repairs (docx4j CR-052 phase 0). Word's check of these saves is the
user's.
"""

from __future__ import annotations

import copy
import io
import logging
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py import create_package, load  # noqa: E402
from docx4j_py.child import (  # noqa: E402
    BoundPrefixes,
    deep_copy,
    iter_tree,
    set_source_bindings,
    source_bindings,
)
from docx4j_py.wml import P, to_xml, wml  # noqa: E402

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XMLNS = "http://www.w3.org/2000/xmlns/"
ZZ9 = "http://example.com/zz9"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
CX1 = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
EXCEL_VML = "urn:schemas-microsoft-com:office:excel"

#: Choices naming zz9 (three) and shape2010 (one), each declared on its mc:AlternateContent.
PROBE = ROOT / "tests" / "fixtures" / "text-extraction" / "phase0" / "extraction-probe.docx"
#: Word 365's waterfall chart: a Choice Requires="cx1" over cx content, cx1 declared on the root only.
WATERFALL = ROOT / "tests" / "fixtures" / "waterfall.docx"


class Found:
    """Each prefix an mc attribute names: where it is declared, or how often it is not."""

    def __init__(self, root: etree._Element) -> None:
        self.root = root
        self.declared_on: dict[str, set[str]] = {}
        self.undeclared: dict[str, int] = {}

    def on(self, prefix: str) -> str:
        """The one kind of element declaring the prefix wherever it is named."""
        kinds = self.declared_on.get(prefix)
        assert kinds is not None, f"{prefix} resolved: {self}"
        assert len(kinds) == 1, f"{prefix} declared on one kind of element: {self}"
        return next(iter(kinds))

    def __repr__(self) -> str:
        return f"declared on {self.declared_on}, undeclared {self.undeclared}"


def _declarer(element: etree._Element, prefix: str) -> etree._Element:
    """The nearest element, from `element` up, declaring the prefix."""
    node: etree._Element | None = element
    while node is not None:
        parent = node.getparent()
        if prefix in node.nsmap and (parent is None or parent.nsmap.get(prefix) != node.nsmap[prefix]):
            return node
        node = parent
    raise AssertionError(f"no declaration of {prefix}")


def check_xml(xml: bytes) -> Found:
    """Every prefix an mc attribute of the part names, resolved where it is named."""
    root = etree.fromstring(xml)
    found = Found(root)
    for element in root.iter():
        if not isinstance(element.tag, str):
            continue
        named: list[str] = []
        if element.tag == f"{{{MC}}}Choice" and element.get("Requires") is not None:
            named.append(element.get("Requires"))
        for local in ("Ignorable", "MustUnderstand"):
            value = element.get(f"{{{MC}}}{local}")
            if value is not None:
                named.append(value)
        for value in named:
            for prefix in value.split():
                if prefix in element.nsmap:
                    found.declared_on.setdefault(prefix, set()).add(etree.QName(_declarer(element, prefix)).localname)
                else:
                    found.undeclared[prefix] = found.undeclared.get(prefix, 0) + 1
    return found


def document_xml(docx: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        return z.read("word/document.xml")


def check(pkg) -> Found:
    return check_xml(document_xml(pkg.save()))


def paragraph_with_choice(choice_attributes: str, declarations: str, content: str) -> str:
    """A paragraph with one mc:AlternateContent, its fragment declaring the prefixes it names."""
    return (
        f'<w:p xmlns:w="{W}" xmlns:mc="{MC}" {declarations}>'
        f"<w:r><mc:AlternateContent><mc:Choice {choice_attributes}>{content}</mc:Choice>"
        "<mc:Fallback><w:t>f</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>"
    )


def choices(tree, want=lambda c: True) -> list:
    """Every typed mc:Choice under `tree`."""
    return [n for n in iter_tree(tree) if type(n).__name__ == "AlternateContentChoice" and want(n)]


def zz9_only(choice) -> bool:
    return choice.requires == "zz9"


def paragraphs(pkg, want=lambda c: True) -> list:
    """The body's paragraphs holding a Choice the predicate selects."""
    return [p for p in pkg.main_document_part.contents.body.content if isinstance(p, P) and choices(p, want)]


def append(pkg, paragraph) -> None:
    pkg.main_document_part.contents.body.content.append(paragraph)


def root_ignorable_zz9(bound: str = ZZ9) -> bytes:
    """A document part whose root declares xmlns:zz9 and names it first in mc:Ignorable."""
    pkg = create_package()
    pkg.main_document_part.contents.body.content.append(wml("<w:p><w:r><w:t>x</w:t></w:r></w:p>"))
    xml = document_xml(pkg.save()).decode("utf-8")
    xml = xml.replace("<w:document ", f'<w:document xmlns:zz9="{bound}" ', 1)
    head = xml[: xml.index(">", xml.index("<w:document"))]
    if "xmlns:mc=" not in head:
        xml = xml.replace("<w:document ", f'<w:document xmlns:mc="{MC}" ', 1)
    if 'mc:Ignorable="' in head:
        return xml.replace('mc:Ignorable="', 'mc:Ignorable="zz9 ', 1).encode("utf-8")
    return xml.replace("<w:document ", '<w:document mc:Ignorable="zz9" ', 1).encode("utf-8")


def with_document(document: bytes) -> bytes:
    """The package create_package makes, its main part's bytes replaced."""
    base = create_package().save()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(base)) as z, zipfile.ZipFile(out, "w") as o:
        for info in z.infolist():
            data = z.read(info.filename)
            if info.filename == "word/document.xml":
                data = document
            o.writestr(info.filename, data)
    return out.getvalue()


# ---------------------------------------------------------------------------
# section 1, case 1: a part saved as loaded
# ---------------------------------------------------------------------------


def test_a_part_nobody_unmarshalled_keeps_every_declaration():
    """Saved from its source bytes: every declaration stays where the source made it."""
    pkg = load(PROBE)
    assert not pkg.main_document_part.is_unmarshalled
    found = check(pkg)
    assert found.undeclared == {}
    assert {"zz9", "shape2010", "wps", "w14"} <= set(found.declared_on)
    assert found.on("zz9") == "AlternateContent"


def test_an_unmarshalled_part_declares_every_prefix_where_word_does():
    """wps and w14 on the root; zz9 (three Choices) and shape2010 on the Choice naming each, ahead of Requires."""
    pkg = load(PROBE)
    pkg.main_document_part.contents  # noqa: B018 - unmarshalled, as any edit does
    xml = document_xml(pkg.save())
    found = check_xml(xml)
    assert found.undeclared == {}
    assert found.on("wps") == "document"
    assert found.on("w14") == "document"
    assert found.on("zz9") == "Choice"
    assert found.on("shape2010") == "Choice"
    assert found.root.nsmap["wps"] == WPS
    assert "zz9" not in found.root.nsmap
    text = xml.decode("utf-8")
    assert f'<mc:Choice xmlns:zz9="{ZZ9}" Requires="zz9">' in text
    assert f'<mc:Choice xmlns:shape2010="{WPS}" Requires="shape2010">' in text
    assert "<wps:wsp>" in text  # the content keeps the table's prefix
    assert text.count("xmlns:zz9=") == 3
    assert "<mc:Fallback xmlns" not in text


# ---------------------------------------------------------------------------
# section 1, case 2: the root's mc:Ignorable naming a private prefix
# ---------------------------------------------------------------------------


def test_a_root_ignorable_prefix_loaded_is_declared_on_the_root():
    """Declared whether saved unread (its bytes) or unmarshalled first (the source's root binding)."""
    source = with_document(root_ignorable_zz9())
    unread = check(load(source))
    assert unread.undeclared == {}
    assert unread.root.nsmap["zz9"] == ZZ9

    pkg = load(source)
    pkg.main_document_part.contents  # noqa: B018
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("zz9") == "document"
    assert found.root.nsmap["zz9"] == ZZ9
    assert found.root.get(f"{{{MC}}}Ignorable").split()[0] == "zz9"


def test_a_prefix_the_root_ignorable_names_is_declared_once_on_the_root():
    """A Choice naming the root's Ignorable prefix, bound alike: one declaration, on the root."""
    xml = root_ignorable_zz9().replace(
        b"</w:body>", paragraph_with_choice('Requires="zz9"', "", "<w:t>c</w:t>").encode() + b"</w:body>"
    )
    pkg = load(with_document(xml))
    pkg.main_document_part.contents  # noqa: B018
    saved = document_xml(pkg.save())
    found = check_xml(saved)
    assert found.undeclared == {}
    assert found.on("zz9") == "document"
    assert saved.count(b"xmlns:zz9=") == 1


# ---------------------------------------------------------------------------
# section 1, case 3: a Choice built in code
# ---------------------------------------------------------------------------


def test_a_choice_built_in_code_naming_a_table_prefix_is_declared_on_the_root():
    """cx1 is the table's, so it goes on the root."""
    pkg = create_package()
    append(pkg, wml(paragraph_with_choice('Requires="cx1"', f'xmlns:cx1="{CX1}"', "<w:t>c</w:t>")))
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("cx1") == "document"
    assert found.root.nsmap["cx1"] == CX1


def test_a_choice_built_in_code_naming_a_private_prefix(caplog):
    """Recorded from its fragment's declaration; renamed in code, undeclared; bound in code, declared."""
    pkg = create_package()
    append(pkg, wml(paragraph_with_choice('Requires="zz9"', f'xmlns:zz9="{ZZ9}"', "<w:t>c</w:t>")))
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("zz9") == "Choice"

    pkg = create_package()
    paragraph = wml(paragraph_with_choice('Requires="zz8"', f'xmlns:zz8="{ZZ9}"', "<w:t>c</w:t>"))
    choice = choices(paragraph)[0]
    choice.requires = "zz9"  # renamed after the parse: the record knew zz8, not zz9
    append(pkg, paragraph)
    with caplog.at_level(logging.WARNING, logger="docx4j_py.openpackaging"):
        found = check(pkg)
    assert found.undeclared == {"zz9": 1}
    assert any("zz9" in record.getMessage() for record in caplog.records)

    set_source_bindings(choice, {"zz9": ZZ9})
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("zz9") == "Choice"


# ---------------------------------------------------------------------------
# section 1, case 4: content moved or copied into another package
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("copied", [False, True], ids=["moved", "copied"])
def test_zz9_moved_or_copied_is_declared_on_the_choice(copied):
    source = load(PROBE)
    dest = create_package()
    found_paragraphs = paragraphs(source, zz9_only)
    assert len(found_paragraphs) == 2
    for paragraph in found_paragraphs:
        if not copied:
            source.main_document_part.contents.body.content.remove(paragraph)
        append(dest, deep_copy(paragraph) if copied else paragraph)
    found = check(dest)
    assert found.undeclared == {}
    assert found.on("zz9") == "Choice"
    first = next(found.root.iter(f"{{{MC}}}Choice"))
    assert _declarer(first, "zz9").nsmap["zz9"] == ZZ9


def test_every_choice_copied_is_declared():
    source = load(PROBE)
    dest = create_package()
    for paragraph in paragraphs(source):
        append(dest, deep_copy(paragraph))
    found = check(dest)
    assert found.undeclared == {}
    assert found.on("wps") == "document"
    assert found.on("w14") == "document"
    assert found.on("zz9") == "Choice"
    assert found.on("shape2010") == "Choice"


# ---------------------------------------------------------------------------
# section 2's Office case: Word's chartex chart
# ---------------------------------------------------------------------------


def test_words_chartex_chart_resaved_keeps_cx1():
    pkg = load(WATERFALL)
    assert len(choices(pkg.main_document_part.contents, lambda c: c.requires == "cx1")) == 1
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("cx1") == "document"
    assert found.root.nsmap["cx1"] == CX1


def test_words_chartex_chart_copied_or_merged_keeps_cx1():
    source = load(WATERFALL)
    dest = create_package()
    for paragraph in paragraphs(source):
        append(dest, deep_copy(paragraph))
    copied = check(dest)
    assert copied.undeclared == {}
    assert copied.on("cx1") == "document"

    source = load(WATERFALL)
    probe = load(PROBE)
    for paragraph in paragraphs(source):
        source.main_document_part.contents.body.content.remove(paragraph)
        append(probe, paragraph)
    merged = check(probe)
    assert merged.undeclared == {}
    assert merged.on("cx1") == "document"
    assert merged.on("wps") == "document"
    assert merged.on("zz9") == "Choice"
    assert merged.on("shape2010") == "Choice"


# ---------------------------------------------------------------------------
# D1 (b): a Choice's own mc:Ignorable and mc:MustUnderstand
# ---------------------------------------------------------------------------


def test_a_choices_own_ignorable_and_must_understand_are_declared():
    pkg = create_package()
    append(
        pkg,
        wml(
            paragraph_with_choice(
                'Requires="wps" mc:Ignorable="zz9" mc:MustUnderstand="w14"',
                f'xmlns:zz9="{ZZ9}" xmlns:wps="{WPS}" xmlns:w14="{W14}"',
                "<w:t>c</w:t>",
            )
        ),
    )
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("wps") == "document"
    assert found.on("w14") == "document"
    assert found.on("zz9") == "Choice"


# ---------------------------------------------------------------------------
# D3: the source's binding first; a clash with the Choice's own content
# ---------------------------------------------------------------------------


def test_a_table_prefix_the_source_bound_otherwise(caplog):
    """x as Excel's VML binds it goes on the Choice; where the content uses the table's x, the content's binding stands."""
    from docx4j_py.namespaces import PREFIXES

    assert PREFIXES["x"] == SML
    pkg = create_package()
    append(pkg, wml(paragraph_with_choice('Requires="x"', f'xmlns:x="{EXCEL_VML}"', "<w:t>c</w:t>")))
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("x") == "Choice"
    choice = next(found.root.iter(f"{{{MC}}}Choice"))
    assert choice.nsmap["x"] == EXCEL_VML

    # docx4j's clash case: the Choice's content in SpreadsheetML. Its marshaller writes that
    # content x:, so a declaration of x on the Choice would rebind it; this serialiser writes
    # SpreadsheetML s:, so x is free and takes the source's binding, the content keeping its own
    pkg = create_package()
    append(
        pkg,
        wml(paragraph_with_choice('Requires="x"', f'xmlns:x="{EXCEL_VML}" xmlns:s="{SML}"', "<s:sheetData/>")),
    )
    found = check(pkg)
    assert found.undeclared == {}
    choice = next(found.root.iter(f"{{{MC}}}Choice"))
    assert choice.nsmap["x"] == EXCEL_VML
    assert etree.QName(choice[0]).namespace == SML
    assert choice[0].prefix == "s"
    del caplog


def test_a_choice_in_a_branch_is_recorded_and_declared():
    """A Choice inside another's branch: recorded; on the root where its Ignorable names it, else on the Choice.

    docx4j's case puts the inner element in a content control directly in the
    branch (its CR-051 section 11.3). Here a ``w:sdt`` in a run-level branch
    is read as the block-level control and loses its runs to the skipped
    report, so the inner element stands in the branch itself.
    """
    inner = (
        f'<mc:AlternateContent xmlns:zz9="{ZZ9}"><mc:Choice Requires="zz9"><w:t>in</w:t></mc:Choice>'
        "<mc:Fallback><w:t>out</w:t></mc:Fallback></mc:AlternateContent>"
    )
    paragraph = paragraph_with_choice('Requires="wps"', f'xmlns:wps="{WPS}"', f"<w:t>s</w:t>{inner}").encode()
    xml = root_ignorable_zz9().replace(b"</w:body>", paragraph + b"</w:body>")
    pkg = load(with_document(xml))
    nested = choices(pkg.main_document_part.contents, zz9_only)
    assert [source_bindings(n) for n in nested] == [{"zz9": ZZ9}]
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("zz9") == "document"

    head, body = xml.split(b"<w:body>", 1)
    head = head.replace(b'mc:Ignorable="zz9 ', b'mc:Ignorable="').replace(b'mc:Ignorable="zz9" ', b"")
    head = head.replace(f'xmlns:zz9="{ZZ9}" '.encode(), b"")
    assert b"zz9" not in head
    pkg = load(with_document(head + b"<w:body>" + body))
    pkg.main_document_part.contents  # noqa: B018
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("zz9") == "Choice"


def test_a_choice_does_not_rebind_a_prefix_its_content_uses_as_the_root_binds_it():
    x, y = "http://example.com/x", "http://example.com/y"
    xml = root_ignorable_zz9(x).replace(
        b"</w:body>", paragraph_with_choice('Requires="zz9"', f'xmlns:zz9="{y}"', "<w:t>c</w:t>").encode() + b"</w:body>"
    )
    pkg = load(with_document(xml))
    pkg.main_document_part.contents  # noqa: B018
    found = check(pkg)
    assert found.undeclared == {}
    assert found.root.nsmap["zz9"] == x
    choice = next(found.root.iter(f"{{{MC}}}Choice"))
    assert choice.nsmap["zz9"] == y  # the Choice's own binding

    # content in the root's namespace, by the prefix the root binds (as content moved in from a
    # document that bound zz9 so would be): a declaration on the Choice would rebind it
    from docx4j_xsdata.formats.dataclass.models.generics import AnyElement

    choice = choices(pkg.main_document_part.contents, zz9_only)[0]
    choice.any_element.clear()
    choice.any_element.append(AnyElement(qname=f"{{{x}}}thing"))
    found = check(pkg)
    assert found.undeclared == {}
    assert found.on("zz9") == "document"
    choice = next(found.root.iter(f"{{{MC}}}Choice"))
    assert choice.nsmap["zz9"] == x  # the content's binding stands
    assert etree.QName(choice[0]).namespace == x


def test_to_xml_declares_on_the_choice():
    """A fragment's marshal follows the part's rule (docx4j's marshalToString)."""
    paragraph = wml(paragraph_with_choice('Requires="zz9"', f'xmlns:zz9="{ZZ9}"', "<w:t>c</w:t>"))
    xml = to_xml(paragraph)
    assert f'<mc:Choice xmlns:zz9="{ZZ9}" Requires="zz9">' in xml
    found = check_xml(xml.encode("utf-8"))
    assert found.undeclared == {}
    assert found.on("zz9") == "Choice"


def test_a_load_records_the_bindings_on_the_choice_and_the_root():
    pkg = load(PROBE)
    found = {str(c.requires): source_bindings(c) for c in choices(pkg.main_document_part.contents)}
    assert found["zz9"] == {"zz9": ZZ9}
    assert found["shape2010"] == {"shape2010": WPS}
    assert found["wps zz9"] == {"wps": WPS, "zz9": ZZ9}
    assert isinstance(choices(pkg.main_document_part.contents, zz9_only)[0].requires, BoundPrefixes)
    root = load(with_document(root_ignorable_zz9())).main_document_part
    root.contents  # noqa: B018
    assert root.source_ns_map["zz9"] == ZZ9


def test_a_deep_copy_carries_the_choices_record():
    pkg = load(PROBE)
    choice = choices(pkg.main_document_part.contents, zz9_only)[0]
    for copied in (deep_copy(choice), copy.deepcopy(choice)):
        assert copied is not choice
        assert source_bindings(copied) == {"zz9": ZZ9}
        assert copied == choice  # the record is not part of equality
