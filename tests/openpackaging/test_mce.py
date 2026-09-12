"""Markup compatibility: the resolved view, the lossless default, the preprocessor.

CR-002 sections 5.6 and 9, decided question 1.

Three things are on trial:

* ``mce="resolve"``, ``"all"`` and ``"none"`` on the traversal functions;
* that the *default* read loses nothing --- both branches of every
  ``mc:AlternateContent`` in ``samples/`` survive a round trip;
* ``LoadOptions(mce_preprocess=True)``, which does what docx4j's
  ``mc-preprocessor.xslt`` does and is opt-in for that reason.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_mce.py -q
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py import namespaces as ns  # noqa: E402
from docx4j_py.child import iter_children, mce_branch  # noqa: E402
from docx4j_py.openpackaging import (  # noqa: E402
    LoadOptions,
    OpcPackage,
    WordprocessingMLPackage,
)
from docx4j_py.openpackaging.mce import (  # noqa: E402
    MCE_NS,
    mce_preprocess_bytes,
    resolve_alternate_content,
)
from docx4j_py.wml import find, text_of, wml  # noqa: E402

#: The samples that hold ``mc:AlternateContent`` somewhere.
MCE_SAMPLES = [
    "2010-glow-then-AlternateContent.docx",
    "2010-mcAlternateContent-in-header.docx",
    "DrawingML_GraphicData_wps.docx",
]

#: Every sample, for the checks that are about the *view* rather than the markup.
ALL_SAMPLES = MCE_SAMPLES + ["2016_image_with_text_effects.docx", "toc.docx"]

AC = f"{{{MCE_NS}}}AlternateContent"
CHOICE = f"{{{MCE_NS}}}Choice"
FALLBACK = f"{{{MCE_NS}}}Fallback"


# ---------------------------------------------------------------------------
# UNDERSTOOD, generated rather than hand kept (CR-002 section 9)
# ---------------------------------------------------------------------------


def test_understood_is_generated_from_the_model():
    from docx4j_py.el_index import EL_MODULES

    understood = ns.UNDERSTOOD
    assert set(understood.values()) == set(EL_MODULES)
    for prefix, uri in understood.items():
        assert ns.PREFIXES[prefix] == uri


def test_understood_names_what_the_corpus_requires():
    """``wps`` is what every ``mc:Choice`` in ``samples/`` requires."""
    assert "wps" in ns.UNDERSTOOD
    assert "w14" in ns.UNDERSTOOD
    assert "wpg" in ns.UNDERSTOOD


def test_understood_is_narrower_than_the_prefix_table():
    """A prefix with no classes behind it must not count as understood."""
    assert set(ns.UNDERSTOOD) < set(ns.PREFIXES)
    assert "wpi" not in ns.UNDERSTOOD  # wordprocessingInk: no bindings in this build
    assert "v" not in ns.UNDERSTOOD  # VML: CR-001 section 13.5, Phase D


# ---------------------------------------------------------------------------
# the three modes on a fragment
# ---------------------------------------------------------------------------


FRAGMENT = """
<w:r>
  <w:t>a</w:t>
  <mc:AlternateContent>
    <mc:Choice Requires="wps"><w:t>-modern</w:t></mc:Choice>
    <mc:Fallback><w:t>-legacy</w:t></mc:Fallback>
  </mc:AlternateContent>
  <mc:AlternateContent>
    <mc:Choice Requires="zz99"><w:t>-unknown</w:t></mc:Choice>
    <mc:Fallback><w:t>-fallback</w:t></mc:Fallback>
  </mc:AlternateContent>
</w:r>
"""


@pytest.mark.parametrize(
    ("mode", "expected"),
    [
        ("resolve", "a-modern-fallback"),
        ("all", "a-modern-legacy-unknown-fallback"),
        ("none", "a"),
    ],
)
def test_text_of_takes_the_branch_the_mode_asks_for(mode, expected):
    """`resolve` is Word's view: the understood choice, else the fallback."""
    assert text_of(wml(FRAGMENT), mce=mode) == expected


def test_resolve_is_the_default():
    assert text_of(wml(FRAGMENT)) == text_of(wml(FRAGMENT), mce="resolve")


def test_the_resolved_view_hides_the_wrappers():
    """docx4j's preprocessor removes the element; the resolved view does too."""
    from docx4j_py.mce import AlternateContent, AlternateContentChoice

    run = wml(FRAGMENT)
    assert find(run, AlternateContent, mce="all")
    assert find(run, AlternateContent, mce="resolve") == []
    assert find(run, AlternateContentChoice, mce="resolve") == []
    assert find(run, AlternateContent, mce="none") == []


def test_mce_branch_picks_the_first_understood_choice():
    from docx4j_py.mce import AlternateContent

    run = wml(FRAGMENT)
    contents = find(run, AlternateContent, mce="all")
    assert len(contents) == 2
    first, second = contents
    assert mce_branch(first).requires == "wps"
    assert mce_branch(second) is second.fallback


def test_a_choice_with_no_fallback_and_nothing_understood_contributes_nothing():
    run = wml(
        "<w:r><w:t>a</w:t><mc:AlternateContent>"
        '<mc:Choice Requires="zz99"><w:t>-x</w:t></mc:Choice>'
        "</mc:AlternateContent></w:r>"
    )
    assert text_of(run, mce="resolve") == "a"
    assert text_of(run, mce="all") == "a-x"


def test_link_parents_and_iter_tree_still_see_everything():
    """CR-002 section 9: ``link_parents`` must keep visiting everything."""
    from docx4j_py.child import iter_tree
    from docx4j_py.mce import AlternateContent

    run = wml(FRAGMENT)
    assert any(isinstance(n, AlternateContent) for n in iter_tree(run))
    names = [q for q, _c in iter_children(run, mce="all")]
    assert names.count(AC) == 2


# ---------------------------------------------------------------------------
# the corpus: nothing is lost by default
# ---------------------------------------------------------------------------


def _counts(data: bytes) -> tuple[int, int, int]:
    root = etree.fromstring(data)
    return (
        len(list(root.iter(AC))),
        len(list(root.iter(CHOICE))),
        len(list(root.iter(FALLBACK))),
    )


@pytest.mark.parametrize("sample", MCE_SAMPLES)
def test_both_branches_survive_a_round_trip(sample):
    """Decided question 1: lossless on read, by default."""
    path = ROOT / "samples" / sample
    with zipfile.ZipFile(path) as zf:
        source = {
            n: zf.read(n)
            for n in zf.namelist()
            if n.endswith(".xml") and b"AlternateContent" in zf.read(n)
        }
    assert source, f"{sample} must hold mc:AlternateContent"

    with OpcPackage.load(path) as pkg:
        for store_name, data in source.items():
            part = pkg.get_part("/" + store_name)
            if part is None or not hasattr(part, "contents"):
                continue
            part.contents  # noqa: B018
            assert _counts(data) == _counts(part.xml), store_name


@pytest.mark.parametrize("sample", ALL_SAMPLES)
def test_text_of_does_not_double_count_a_text_box(sample):
    """The view Word takes, over the real corpus.

    ``mce="resolve"`` must never report a string twice because the document
    wrote it once as a ``wps`` shape and once as a VML fallback. Over these
    fixtures the resolved view is a *subset* of the ``all`` view and drops
    nothing a reader would see.
    """
    with WordprocessingMLPackage.load(ROOT / "samples" / sample) as pkg:
        document = pkg.main_document_part.contents
        resolved = text_of(document, mce="resolve")
        every = text_of(document, mce="all")

    assert len(resolved) <= len(every)
    for line in resolved.splitlines():
        if line.strip():
            assert every.count(line) >= resolved.count(line)
    # and no line is repeated by the resolution itself
    lines = [line for line in resolved.splitlines() if line.strip()]
    for line in lines:
        assert resolved.count(line) <= every.count(line)


def test_the_wps_text_box_is_reported_once():
    """``DrawingML_GraphicData_wps.docx``: one shape, two branches, one text."""
    path = ROOT / "samples" / "DrawingML_GraphicData_wps.docx"
    with zipfile.ZipFile(path) as zf:
        data = zf.read("word/document.xml")
    root = etree.fromstring(data)
    ac = next(root.iter(AC))
    choice = next(c for c in ac if c.tag == CHOICE)
    fallback = next(c for c in ac if c.tag == FALLBACK)
    w_t = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
    in_choice = [t.text for t in choice.iter(w_t)]
    in_fallback = [t.text for t in fallback.iter(w_t)]
    assert in_choice == in_fallback, "the fixture must say the same thing twice"

    with WordprocessingMLPackage.load(path) as pkg:
        from docx4j_py.mce import AlternateContent
        from docx4j_py.wml import find

        document = pkg.main_document_part.contents
        ac = find(document, AlternateContent, mce="all")[0]
        # on the AlternateContent itself the view is observable: once resolved, twice with all
        assert text_of(ac, mce="resolve").count(in_choice[0]) == 1
        assert text_of(ac, mce="all").count(in_choice[0]) == 2
        # at document level the box is not reached from its run in any mode: CR-002 12.6,
        # open question 7. Pinned so a change is a decision, not an accident.
        assert text_of(document, mce="all").count(in_choice[0]) == 0


# ---------------------------------------------------------------------------
# the preprocessor, LoadOptions(mce_preprocess=True)
# ---------------------------------------------------------------------------


def test_the_preprocessor_replaces_alternate_content_with_the_chosen_branch():
    data = (
        b'<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        b' xmlns:mc="' + MCE_NS.encode() + b'"'
        b' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        b"<w:p/>"
        b'<mc:AlternateContent><mc:Choice Requires="wps"><w:tbl/></mc:Choice>'
        b"<mc:Fallback><w:sdt/></mc:Fallback></mc:AlternateContent>"
        b"</w:body>"
    )
    out = etree.fromstring(mce_preprocess_bytes(data))
    tags = [etree.QName(c).localname for c in out]
    assert tags == ["p", "tbl"]
    assert list(out.iter(AC)) == []


def test_the_preprocessor_falls_back_when_nothing_is_understood():
    data = (
        b'<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        b' xmlns:mc="' + MCE_NS.encode() + b'" xmlns:zz="urn:x">'
        b'<mc:AlternateContent><mc:Choice Requires="zz"><w:tbl/></mc:Choice>'
        b"<mc:Fallback><w:sdt/></mc:Fallback></mc:AlternateContent>"
        b"</w:body>"
    )
    out = etree.fromstring(mce_preprocess_bytes(data))
    assert [etree.QName(c).localname for c in out] == ["sdt"]


def test_the_preprocessor_drops_what_has_no_fallback():
    data = (
        b'<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        b' xmlns:mc="' + MCE_NS.encode() + b'" xmlns:zz="urn:x">'
        b'<mc:AlternateContent><mc:Choice Requires="zz"><w:tbl/></mc:Choice>'
        b"</mc:AlternateContent><w:p/></w:body>"
    )
    out = etree.fromstring(mce_preprocess_bytes(data))
    assert [etree.QName(c).localname for c in out] == ["p"]


def test_the_preprocessor_resolves_innermost_first():
    data = (
        b'<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        b' xmlns:mc="' + MCE_NS.encode() + b'"'
        b' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        b'<mc:AlternateContent><mc:Choice Requires="wps">'
        b'<mc:AlternateContent><mc:Choice Requires="wps"><w:tbl/></mc:Choice>'
        b"<mc:Fallback><w:sdt/></mc:Fallback></mc:AlternateContent>"
        b"</mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent></w:body>"
    )
    out = etree.fromstring(mce_preprocess_bytes(data))
    assert [etree.QName(c).localname for c in out] == ["tbl"]
    assert list(out.iter(AC)) == []


@pytest.mark.parametrize("sample", ALL_SAMPLES)
def test_mce_preprocess_on_load_gives_a_tree_with_no_alternate_content(sample):
    path = ROOT / "samples" / sample
    with WordprocessingMLPackage.load(path, options=LoadOptions(mce_preprocess=True)) as pkg:
        for part in pkg.parts.parts():
            if not hasattr(part, "contents") or part.model_class is None:
                continue
            part.contents  # noqa: B018
            assert list(etree.fromstring(part.xml).iter(AC)) == [], part.part_name


@pytest.mark.parametrize("sample", ALL_SAMPLES)
def test_preprocessing_gives_the_same_text_as_the_resolved_view(sample):
    """The two routes to Word's view must agree."""
    path = ROOT / "samples" / sample
    with WordprocessingMLPackage.load(path) as lossless:
        a = text_of(lossless.main_document_part.contents, mce="resolve")
    with WordprocessingMLPackage.load(
        path, options=LoadOptions(mce_preprocess=True)
    ) as preprocessed:
        b = text_of(preprocessed.main_document_part.contents, mce="all")
    assert a == b


def test_resolve_alternate_content_counts_what_it_did():
    data = (
        b'<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        b' xmlns:mc="' + MCE_NS.encode() + b'" xmlns:zz="urn:x">'
        b"<mc:AlternateContent><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>"
        b"<mc:AlternateContent><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>"
        b"</w:body>"
    )
    root = etree.fromstring(data)
    assert resolve_alternate_content(root) == 2
