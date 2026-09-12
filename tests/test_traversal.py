"""``walk``, ``find`` and ``text_of``, CR-001 section 6.2.

The parity test is the one that matters: ``text_of`` over every
``document.xml`` in ``samples/`` against an independent extraction with lxml
that follows the same docx4j ``TextUtils`` rules. One implementation reads the
typed tree through the class metadata, the other reads the XML; if they agree
on ten real documents the metadata-driven traversal is enumerating the same
children the document holds.

    .venv-fork/bin/python -m pytest tests/test_traversal.py -q
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py import wml as W  # noqa: E402
from docx4j_py.child import iter_children  # noqa: E402
from docx4j_py.namespaces import WML_NS  # noqa: E402
from docx4j_py.runtime import parser  # noqa: E402
from docx4j_py.traversal import (  # noqa: E402
    element_name,
    find,
    iter_nodes,
    text_of,
    walk,
)
from docx4j_py.wml import el, p, r, wml  # noqa: E402

SAMPLES = sorted((ROOT / "samples").glob("*.docx"))


# ---------------------------------------------------------------------------
# walk and find
# ---------------------------------------------------------------------------


def test_walk_is_pre_order_and_reports_the_element_name():
    para = wml('<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>')
    seen: list[tuple[str, str | None]] = []
    walk(para, lambda node, parent, name: seen.append((type(node).__name__, name)))
    assert seen[0] == ("P", f"{{{WML_NS}}}p")
    assert ("PPr", f"{{{WML_NS}}}pPr") in seen
    assert ("R", f"{{{WML_NS}}}r") in seen
    assert ("RT", f"{{{WML_NS}}}t") in seen
    # pre-order: a node comes before its children
    assert seen.index(("R", f"{{{WML_NS}}}r")) < seen.index(("RT", f"{{{WML_NS}}}t"))


def test_walk_parents_are_the_containing_nodes():
    para = wml("<w:p><w:r><w:t>Hi</w:t></w:r></w:p>")
    parents = {}
    walk(
        para,
        lambda node, parent, name: parents.__setitem__(type(node).__name__, parent),
    )
    assert parents["P"] is None
    assert parents["R"] is para
    assert parents["RT"] is para.content[0]


def test_visitor_returning_false_stops_descending():
    para = wml("<w:p><w:r><w:t>Hi</w:t></w:r></w:p>")
    seen = []

    def visitor(node, parent, name):
        seen.append(type(node).__name__)
        return not isinstance(node, W.R)

    walk(para, visitor)
    assert seen == ["P", "R"]


def test_iter_nodes_is_the_same_traversal():
    para = wml("<w:p><w:r><w:t>Hi</w:t></w:r></w:p>")
    by_walk: list = []
    walk(para, lambda node, parent, name: by_walk.append(node))
    assert [id(n) for n in iter_nodes(para)] == [id(n) for n in by_walk]


def test_find_matches_subclasses():
    """Decided question 2: `find(root, CTTrackChange)` must match RunIns."""
    body = wml(
        "<w:body>"
        '<w:p><w:ins w:id="1"><w:r><w:t>new</w:t></w:r></w:ins>'
        '<w:del w:id="2"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>'
        "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"
        "</w:body>"
    )
    assert len(find(body, W.P)) == 2
    assert len(find(body, W.Tbl)) == 1
    assert len(find(body, W.RunIns)) == 1
    assert len(find(body, W.RunDel)) == 1
    assert len(find(body, W.CTTrackChange)) == 2  # the base of both
    assert len(find(body, (W.P, W.Tbl))) == 3
    # `RT` (w:t in a run) is a subclass of `Text`; `DelText` is *not* --- xsdata
    # generated the global element `w:delText` as its own class rather than as a
    # subclass of CT_Text, so `find(root, Text)` does not reach it (CR-001
    # section 14). `(Text, DelText)` is how to ask for both.
    assert len(find(body, W.Text)) == 1
    assert W.Text not in W.DelText.__mro__
    assert len(find(body, (W.Text, W.DelText))) == 2
    assert find(body, W.Body) == [body]


def test_find_on_a_built_tree():
    tree = el.body(content=[p("one"), p("two"), p("three")])
    assert len(find(tree, W.P)) == 3
    assert len(find(tree, W.R)) == 3


def test_element_name_of_the_hard_cases():
    assert element_name(el.p()) == f"{{{WML_NS}}}p"  # a global element
    assert element_name(el.document()) == f"{{{WML_NS}}}document"
    assert element_name(W.Tbl()) == f"{{{WML_NS}}}tbl"  # a type class
    assert element_name(el.t("x")) == f"{{{WML_NS}}}t"  # an intermediate class
    assert element_name(W.TblWidth()) is None  # eight element names, no answer


def test_iter_children_is_the_shared_enumeration():
    para = wml("<w:p><w:pPr/><w:r><w:t>Hi</w:t></w:r></w:p>")
    names = [name for name, _child in iter_children(para)]
    assert names == [f"{{{WML_NS}}}pPr", f"{{{WML_NS}}}r"]


# ---------------------------------------------------------------------------
# text_of
# ---------------------------------------------------------------------------


def test_text_of_the_readme_example():
    assert text_of(p("Hello World")) == "Hello World"
    assert text_of(r("a\tb\nc")) == "a\tb\nc"


def test_text_of_joins_paragraphs_with_a_newline():
    body = wml(
        "<w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body>"
    )
    assert text_of(body) == "one\ntwo"


def test_text_of_reads_through_the_run_containers_and_skips_deletions():
    para = wml(
        "<w:p>"
        "<w:r><w:t>a</w:t></w:r>"
        '<w:hyperlink r:id="rId1"><w:r><w:t>b</w:t></w:r></w:hyperlink>'
        '<w:ins w:id="1"><w:r><w:t>c</w:t></w:r></w:ins>'
        '<w:del w:id="2"><w:r><w:delText>GONE</w:delText></w:r></w:del>'
        '<w:fldSimple w:instr=" PAGE "><w:r><w:t>d</w:t></w:r></w:fldSimple>'
        "<w:r><w:instrText> REF x </w:instrText></w:r>"
        "</w:p>"
    )
    assert text_of(para) == "abcd"


def test_text_of_the_run_level_characters():
    run = wml(
        "<w:r><w:t>a</w:t><w:tab/><w:br/><w:cr/><w:noBreakHyphen/><w:softHyphen/>"
        '<w:sym w:font="Wingdings" w:char="0041"/></w:r>'
    )
    assert text_of(run) == "a\t\n\n‑­A"


def test_text_of_a_content_control_and_a_table():
    body = wml(
        "<w:body>"
        "<w:sdt><w:sdtContent><w:p><w:r><w:t>controlled</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        "</w:body>"
    )
    assert text_of(body) == "controlled\ncell"


# ---------------------------------------------------------------------------
# parity with an independent lxml extraction
# ---------------------------------------------------------------------------

RUN_TEXT = {
    "tab": "\t",
    "br": "\n",
    "cr": "\n",
    "noBreakHyphen": "‑",
    "softHyphen": "­",
}
EXCLUDED = {"delText", "instrText", "delInstrText", "del"}
RUN_CONTAINERS = {
    "hyperlink",
    "sdt",
    "sdtContent",
    "smartTag",
    "customXml",
    "ins",
    "moveTo",
    "dir",
    "bdo",
    "fldSimple",
}


def _local(node) -> str | None:
    """The WML local name of an element, or None for anything else."""
    tag = node.tag
    if not isinstance(tag, str) or not tag.startswith(f"{{{WML_NS}}}"):
        return None
    return tag.split("}", 1)[1]


def lxml_run_text(node) -> str:
    out = []
    for child in node:
        local = _local(child)
        if local == "t":
            out.append(child.text or "")
        elif local == "sym":
            value = child.get(f"{{{WML_NS}}}char")
            out.append(chr(int(value, 16)) if value and len(value) == 4 else "")
        elif local in RUN_TEXT:
            out.append(RUN_TEXT[local])
    return "".join(out)


def lxml_inline_text(node) -> str:
    out = []
    for child in node:
        local = _local(child)
        if local in EXCLUDED or local is None:
            continue
        if local == "r":
            out.append(lxml_run_text(child))
        elif local in RUN_CONTAINERS or local == "p":
            out.append(lxml_inline_text(child))
    return "".join(out)


def lxml_block_texts(node, out: list[str]) -> None:
    local = _local(node)
    if local in EXCLUDED:
        return
    if local == "r":
        out.append(lxml_run_text(node))
        return
    if local == "p" or local in RUN_CONTAINERS:
        out.append(lxml_inline_text(node))
        return
    for child in node:
        if isinstance(child.tag, str):
            lxml_block_texts(child, out)


def lxml_text_of(xml: bytes) -> str:
    root = etree.fromstring(xml)
    out: list[str] = []
    lxml_block_texts(root, out)
    return "\n".join(out)


@pytest.mark.parametrize("docx", SAMPLES, ids=lambda p: p.name)
def test_text_of_matches_an_lxml_extraction(docx):
    with zipfile.ZipFile(docx) as zf:
        data = zf.read("word/document.xml")
    document = parser().from_bytes(data, W.Document)
    assert text_of(document) == lxml_text_of(data)


def test_the_parity_check_is_not_vacuous():
    """At least one sample has real text, so the comparison means something."""
    with zipfile.ZipFile(ROOT / "samples" / "tables.docx") as zf:
        data = zf.read("word/document.xml")
    assert len(lxml_text_of(data)) > 50
