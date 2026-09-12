"""``wml(...)`` and ``to_xml(...)``, CR-001 section 6.2.

The round-trip contract: ``to_xml(wml(x))`` is canonically equal to ``x`` for
fragments that declare no namespaces of their own, including ``w14`` attributes
and ``mc:AlternateContent``. "Canonically" is ``scripts/canon.py``'s normal form
--- prefix-insensitive, attribute-sorted, whitespace-normalised except under
``xml:space="preserve"`` --- which is the same verdict the corpus round trip
uses.

    .venv-fork/bin/python -m pytest tests/test_fragments.py -q
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon  # noqa: E402

from docx4j_py import wml as W  # noqa: E402
from docx4j_py.fragments import FragmentError, to_xml, wml  # noqa: E402
from docx4j_py.namespaces import MC_IGNORABLE_PREFIXES, PREFIXES  # noqa: E402

# A dozen fragments, none of which declares a namespace: `wml` has to supply
# docx4j's whole prefix table for them to parse at all.
FRAGMENTS = [
    "<w:p/>",
    "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>",
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
    "<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Hello</w:t></w:r></w:p>",
    '<w:p><w:r><w:t xml:space="preserve">  spaced  </w:t></w:r></w:p>',
    '<w:p w14:paraId="0A1B2C3D" w14:textId="77777777" w:rsidR="00AB00AB">'
    "<w:r><w:t>w14</w:t></w:r></w:p>",
    '<w:pPr><w:spacing w:before="240" w:after="120"/><w:jc w:val="both"/></w:pPr>',
    '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>a</w:t><w:tab/><w:br/><w:cr/>'
    '<w:sym w:font="Wingdings" w:char="F0E0"/></w:r>',
    '<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>'
    "<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>",
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>',
    '<w:hyperlink r:id="rId7" w:history="1"><w:r><w:t>link</w:t></w:r></w:hyperlink>',
    '<w:ins w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z">'
    "<w:r><w:t>new</w:t></w:r></w:ins>",
    '<w:sdt><w:sdtPr><w:alias w:val="a"/><w:tag w:val="t"/><w:id w:val="1"/></w:sdtPr>'
    "<w:sdtContent><w:p><w:r><w:t>in</w:t></w:r></w:p></w:sdtContent></w:sdt>",
    "<w:r><mc:AlternateContent>"
    '<mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>'
    '<wp:docPr id="1" name="s"/><a:graphic><a:graphicData uri="u"/></a:graphic>'
    "</wp:inline></w:drawing></mc:Choice>"
    "<mc:Fallback><w:pict/></mc:Fallback>"
    "</mc:AlternateContent></w:r>",
    "<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>",
]


@pytest.mark.parametrize("fragment", FRAGMENTS, ids=range(len(FRAGMENTS)))
def test_round_trip_is_canonically_identical(fragment):
    out = to_xml(wml(fragment))
    report = canon.DiffReport()
    canon.diff_trees(
        canon.normalize_tree(_declare(fragment).encode()),
        canon.normalize_tree(out.encode()),
        report,
    )
    assert report.differences == {}, (fragment, out, report.diffs[:5])


def _declare(fragment: str) -> str:
    """The fragment with every prefix declared, so lxml can parse it alone."""
    from docx4j_py.namespaces import declarations

    head, _, rest = fragment.partition(">")
    closing = "/>" if head.endswith("/") else ">"
    head = head.removesuffix("/")
    return f"{head} {declarations()}{closing}{rest}"


# ---------------------------------------------------------------------------
# what wml returns
# ---------------------------------------------------------------------------


def test_the_class_comes_from_the_el_table():
    assert isinstance(wml("<w:p/>"), W.P)
    assert isinstance(wml("<w:tbl/>"), W.Tbl)
    assert isinstance(wml("<w:r/>"), W.R)
    assert isinstance(wml("<w:pPr/>"), W.PPr)
    assert isinstance(wml("<w:tc/>"), W.Tc)
    assert isinstance(wml("<w:document><w:body/></w:document>"), W.Document)


def test_parents_are_linked():
    para = wml("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>")
    run = para.content[0]
    assert run.parent is para
    assert run.content[0].parent is run
    assert para.parent is None


def test_wrapper_picks_the_scope_for_an_ambiguous_name():
    """``w:sdt`` is four classes; the container decides which.

    And the choice is not cosmetic: a run-level control read as the block-level
    class loses its runs, which strict parsing refuses rather than swallows ---
    the whole point of the skipped-content report.
    """
    block = (
        "<w:sdt><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>"
    )
    run = "<w:sdt><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>"
    row = "<w:sdt><w:sdtContent><w:tr/></w:sdtContent></w:sdt>"
    cell = "<w:sdt><w:sdtContent><w:tc/></w:sdtContent></w:sdt>"

    assert isinstance(wml(block), W.SdtBlock)  # the el table's default
    assert isinstance(wml(block, wrapper="body"), W.SdtBlock)
    assert isinstance(wml(run, wrapper="p"), W.SdtRun)
    assert isinstance(wml(row, wrapper="tbl"), W.CTSdtRow)
    assert isinstance(wml(cell, wrapper="tr"), W.CTSdtCell)
    assert isinstance(wml(block, wrapper=W.Body), W.SdtBlock)

    with pytest.raises(FragmentError, match="skipped"):
        wml(run)  # SdtBlock cannot hold a run


def test_all_returns_the_siblings():
    paragraphs = wml.all("<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p/><w:tbl/>")
    assert [type(o).__name__ for o in paragraphs] == ["P", "P", "Tbl"]


def test_one_element_only_unless_all_is_used():
    with pytest.raises(FragmentError, match="use wml.all"):
        wml("<w:p/><w:p/>")


@pytest.mark.parametrize(
    ("bad", "message"),
    [
        ("", "empty fragment"),
        ("not xml", "holds no element"),
        ("<w:p>", "not well-formed"),
        ("<!-- just a comment -->", "holds no element"),
        ("<w:nosuch/>", "no class is bound"),
    ],
)
def test_bad_fragments_raise(bad, message):
    with pytest.raises(FragmentError, match=message):
        wml(bad)


def test_content_the_bindings_do_not_know_raises_unless_lenient():
    unknown = '<w:p xmlns:zz="urn:zz"><zz:thing/></w:p>'
    with pytest.raises(FragmentError, match="skipped"):
        wml(unknown)
    assert isinstance(wml(unknown, lenient=True), W.P)


def test_declarations_inside_the_fragment_win():
    """A fragment may bind `w` itself; the synthetic root does not override it."""
    other = (
        '<x:p xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )
    assert isinstance(wml(other), W.P)


def test_an_xml_declaration_is_tolerated():
    assert isinstance(wml('<?xml version="1.0" encoding="UTF-8"?><w:p/>'), W.P)


# ---------------------------------------------------------------------------
# to_xml
# ---------------------------------------------------------------------------


def test_to_xml_declares_only_what_the_fragment_uses():
    out = to_xml(wml("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>"))
    assert out.count("xmlns:") == 1
    assert (
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' in out
    )


def test_to_xml_keeps_the_prefixes_mc_ignorable_names():
    """Trimming must not strip a declaration `mc:Ignorable` depends on."""
    out = to_xml(
        wml('<w:document mc:Ignorable="w14 w15"><w:body><w:p/></w:body></w:document>')
    )
    assert 'mc:Ignorable="w14 w15"' in out
    assert "xmlns:w14=" in out and "xmlns:w15=" in out
    assert {"w14", "w15"} <= MC_IGNORABLE_PREFIXES


def test_to_xml_names_a_class_that_has_no_element_name_of_its_own():
    from docx4j_py.wml import el

    assert to_xml(el.t("x")).startswith("<w:t ")
    assert to_xml(el.tab()).startswith("<w:tab ")
    assert to_xml(el.instrText(" PAGE ")).startswith("<w:instrText ")


def test_to_xml_pretty_and_forced_name():
    out = to_xml(wml("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>"), pretty=True)
    assert "\n" in out
    assert to_xml(W.P(), name="{%s}other" % PREFIXES["w"]).startswith("<w:other ")


def test_booleans_are_numeric_as_word_writes_them():
    out = to_xml(wml('<w:rPr><w:b w:val="true"/></w:rPr>'))
    assert 'w:val="1"' in out
