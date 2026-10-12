"""``Range.hyperlink``: Office JS's read and write of a link. CR-003 section 21.

The port of docx4j-core-ts CR-002 sections 20 and 25 (its ``d64ab67`` and the
Word-measured corrections after it): its three tests of section 20 first, then
the markup of its Word checks 15, 16 and 19 --- each case held to the paragraph
XML Word wrote, less Word's own bookkeeping (``w:rsid*``, the ``w14`` ids, the
relationship id) --- then what this repository adds: every link in the corpus
read back, a header's link related from the header, the refusals made before
anything changes, the ``ChangeReport``, ``dry_run``, change tracking and
``to_api_script``.
"""

from __future__ import annotations

import re
import zipfile
from io import BytesIO
from pathlib import Path

import pytest
from conftest import FIXTURES, SAMPLES, fixture, load, part_bytes, reloaded, sample
from lxml import etree

from docx4j_py import create_package
from docx4j_py.model.content import Author, ContentError, Range, SpanError
from docx4j_py.model.content.addresses import package_bodies
from docx4j_py.wml import el, r

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def entry(data: bytes, name: str) -> str:
    """One entry of a saved package, as text."""
    with zipfile.ZipFile(BytesIO(data)) as archive:
        return archive.read(name).decode("utf-8")


def comparable(xml: str) -> bytes:
    """A paragraph's XML in a form Word's can be compared with.

    Word's bookkeeping out (``w:rsid*``, the ``w14`` attributes), the
    relationship id blanked, and the attribute order and namespace declarations
    normalised by exclusive canonicalisation --- the serialiser writes
    ``w:history`` before ``r:id`` where Word writes it after.
    """
    if "xmlns:w=" not in xml:
        xml = xml.replace("<w:p>", f'<w:p xmlns:w="{W_NS}" xmlns:r="{R_NS}">', 1)
    root = etree.fromstring(xml.encode("utf-8"))
    for node in root.iter():
        for name in list(node.attrib):
            local = etree.QName(name).localname
            if name.startswith(f"{{{W14_NS}}}") or local.startswith("rsid"):
                del node.attrib[name]
            elif name == f"{{{R_NS}}}id":
                node.attrib[name] = ""
    return etree.tostring(root, method="c14n", exclusive=True)


def last(pkg):
    """The body's last paragraph."""
    return pkg.body.paragraphs[-1]


def assert_parents(paragraph) -> None:
    """Every node below the paragraph hangs from the node whose list holds it."""
    from docx4j_py.traversal import run_items_of

    def visit(holder) -> None:
        for item in run_items_of(holder) or ():
            assert item.parent is holder, (type(item).__name__, type(holder).__name__)
            if run_items_of(item) is not None:
                visit(item)
            for child in getattr(item, "content", None) or ():
                if hasattr(child, "parent") and run_items_of(item) is None:
                    assert child.parent is item, type(child).__name__

    visit(paragraph.element)


# ---------------------------------------------------------------------------
# docx4j-core-ts CR-002 section 20: its three tests
# ---------------------------------------------------------------------------


def test_an_external_link_over_a_span_with_its_relationship():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("see the docs here")
    assert paragraph.get_range().hyperlink == "", "no hyperlink to begin with"

    span = paragraph.search("the docs")[0]
    span.hyperlink = "https://docx4j.org/"
    assert span.hyperlink == "https://docx4j.org/"
    assert paragraph.text == "see the docs here", "the text is untouched"

    data = pkg.save()
    xml = entry(data, "word/document.xml")
    found = re.search(r'<w:hyperlink[^>]*r:id="([^"]+)"', xml)
    assert found, "a w:hyperlink with an r:id is written"
    # w:history, as Word and docx4j write it; numeric, as every boolean here is
    assert re.search(r'<w:hyperlink[^>]*w:history="1"', xml)
    rels = entry(data, "word/_rels/document.xml.rels")
    assert f'Id="{found.group(1)}"' in rels
    assert 'Target="https://docx4j.org/"' in rels
    assert 'TargetMode="External"' in rels

    # the runs the link covers are exactly the span's
    again = reloaded(pkg).body.paragraphs[0]
    assert again.search("the docs")[0].hyperlink == "https://docx4j.org/"
    assert again.search("here")[0].hyperlink == "", "outside the span there is none"
    assert again.search("see")[0].hyperlink == ""


def test_a_location_inside_the_document_and_removal():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("jump to chapter one")
    span = paragraph.get_range()
    rels = pkg.main_document_part.relationships_part
    count = len(rels)

    span.hyperlink = "#chapter1"
    assert span.hyperlink == "#chapter1", "an anchor, with no relationship"
    assert len(rels) == count
    xml = entry(pkg.save(), "word/document.xml")
    assert re.search(r'<w:hyperlink[^>]*w:anchor="chapter1"', xml)
    assert not re.search(r"<w:hyperlink[^>]*r:id=", xml), "and no r:id"
    assert re.search(r'<w:hyperlink[^>]*w:history="1"', xml), "an internal link carries it too"

    # address and location together
    span.hyperlink = "https://example.com/doc#section2"
    assert span.hyperlink == "https://example.com/doc#section2"
    assert len(rels) == count + 1

    span.hyperlink = ""
    assert span.hyperlink == ""
    assert paragraph.text == "jump to chapter one", "the runs survive the unwrapping"
    after = entry(pkg.save(), "word/document.xml")
    assert "<w:hyperlink" not in after, "and the holder is gone"
    # the relationship stays, as docx4j-core-ts leaves it: harmless, and possibly shared
    assert len(rels) == count + 1


def test_none_removes_as_the_empty_string_does():
    pkg = create_package()
    span = pkg.body.insert_paragraph("a link").get_range()
    span.hyperlink = "#top"
    span.hyperlink = None
    assert span.hyperlink == ""
    assert "<w:hyperlink" not in span.paragraph.get_xml()


def test_reads_what_word_wrote():
    pkg = fixture("hyperlink_dupe.docx")
    links = [p.get_range().hyperlink for p in pkg.body.paragraphs]
    links = [link for link in links if link]
    assert links, "the fixture has hyperlinks"
    for link in links:
        assert re.match(r"^https?://|^#", link), link


# ---------------------------------------------------------------------------
# every link in the corpus reads back
# ---------------------------------------------------------------------------


def _expected(link, part) -> str:
    """What the markup says a link is, read without the getter."""
    anchor = link.anchor or ""
    address = ""
    if link.id:
        rel = part.relationships_part.get_relationship_by_id(link.id)
        address = rel.target if rel is not None else ""
    return f"{address}#{anchor}" if anchor else address


def _link_spans(paragraph):
    """Each w:hyperlink holding text, with the offsets of its text, found from the runs."""
    spans: dict[int, list] = {}
    for segment in paragraph.segments():
        node = segment.run.parent
        while node is not None and node is not paragraph.element:
            if isinstance(node, type(el.hyperlink())):
                spans.setdefault(id(node), [node, segment.start, segment.end])[2] = segment.end
            node = node.parent
    return list(spans.values())


def _corpus() -> list[Path]:
    return sorted([*SAMPLES.glob("*.docx"), *FIXTURES.glob("*.docx")])


def test_every_link_in_the_corpus_reads_back_its_target_or_anchor():
    checked = 0
    documents = set()
    for path in _corpus():
        pkg = load(path)
        for body in package_bodies(pkg):
            for paragraph in body.iter_paragraphs():
                for link, start, end in _link_spans(paragraph):
                    expected = _expected(link, body.part)
                    assert Range(paragraph, start, end).hyperlink == expected, (path.name, start)
                    # a caret inside it, and one at its end, read it too (check 19)
                    caret = (start + end) // 2
                    assert Range(paragraph, caret, caret).hyperlink == expected
                    assert Range(paragraph, end, end).hyperlink == expected
                    checked += 1
                    documents.add(path.name)
    # sample-docx, toc, comments-modern, hyperlink, hyperlink_dupe, loadAndSave, tracked-changes
    assert checked >= 20, checked
    assert len(documents) == 7, sorted(documents)


def test_the_markdown_exporter_and_the_getter_agree_on_a_link_the_setter_made():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("see the docs here, or chapter one")
    paragraph.search("the docs")[0].hyperlink = "https://docx4j.org/"
    paragraph.search("chapter one")[0].hyperlink = "#chapter1"
    assert paragraph.to_markdown() == (
        "see [the docs](https://docx4j.org/) here, or [chapter one](#chapter1)"
    )



def test_the_python_docx_facade_reads_a_link_the_setter_made():
    from docx4j_py.docx import Document

    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("see the docs here, or chapter one")
    paragraph.search("the docs")[0].hyperlink = "https://docx4j.org/"
    paragraph.search("chapter one")[0].hyperlink = "#chapter1"
    links = Document(BytesIO(pkg.save())).paragraphs[0].hyperlinks
    assert [(h.text, h.address, h.fragment) for h in links] == [
        ("the docs", "https://docx4j.org/", ""),
        ("chapter one", "", "chapter1"),
    ]

# ---------------------------------------------------------------------------
# docx4j-core-ts check 15: a link is styled as Word styles one
# ---------------------------------------------------------------------------

WORD_CHECK_15 = {
    "A": '<w:p><w:r><w:t xml:space="preserve">Case A: an external link on </w:t></w:r><w:hyperlink r:id="" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>this text</w:t></w:r></w:hyperlink></w:p>',
    "B": '<w:p><w:r><w:t xml:space="preserve">Case B: an internal link on </w:t></w:r><w:hyperlink w:anchor="_top" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>this text</w:t></w:r></w:hyperlink></w:p>',
    "C": '<w:p><w:r><w:t xml:space="preserve">Case C: a link on </w:t></w:r><w:hyperlink r:id="" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>these strong words</w:t></w:r></w:hyperlink></w:p>',
    "D": '<w:p><w:r><w:t xml:space="preserve">Case D: a link on </w:t></w:r><w:hyperlink r:id="" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/></w:rPr><w:t>this bold red text</w:t></w:r></w:hyperlink></w:p>',
    "E": '<w:p><w:r><w:t xml:space="preserve">Case E: a link set and then </w:t></w:r><w:r><w:t>removed here</w:t></w:r></w:p>',
}


def test_a_link_is_styled_as_word_styles_one_check_15():
    pkg = create_package()
    body = pkg.body
    body.insert_paragraph("Case A: an external link on this text")
    last(pkg).search("this text")[0].hyperlink = "https://example.com/a"
    body.insert_paragraph("Case B: an internal link on this text")
    last(pkg).search("this text")[0].hyperlink = "#_top"
    # C: a run that already names a character style --- Word replaces it
    body.add_object(
        el.p(content=[r("Case C: a link on "), r("these strong words", style="Strong")])
    )
    last(pkg).search("these strong words")[0].hyperlink = "https://example.com/c"
    # D: Word keeps the bold and drops the colour. The run starts as Word's did:
    # Office JS's font.bold wrote w:b alone, where the builder writes w:bCs too
    bold_red = r("this bold red text", bold=True, color="#FF0000")
    bold_red.r_pr.b_cs = None
    body.add_object(el.p(content=[r("Case D: a link on "), bold_red]))
    last(pkg).search("this bold red text")[0].hyperlink = "https://example.com/d"
    # E: set and removed --- the style goes with the link
    body.insert_paragraph("Case E: a link set and then removed here")
    last(pkg).search("removed here")[0].hyperlink = "https://example.com/e"
    last(pkg).search("removed here")[0].hyperlink = ""

    cases = {
        p.text[5]: comparable(p.get_xml()) for p in body.paragraphs if p.text.startswith("Case ")
    }
    for case, xml in WORD_CHECK_15.items():
        assert cases[case] == comparable(xml), case


def test_the_hyperlink_style_is_defined_where_the_document_lacks_it():
    pkg = sample(
        "2010-sample1.docx"
    )  # defines no Hyperlink style, as Word adds one with the first link
    styles = pkg.style_definitions_part
    assert b'w:styleId="Hyperlink"' not in styles.get_xml().encode()
    pkg.body.paragraphs[0].search("first")[0].hyperlink = "https://example.com/"
    assert "/word/styles.xml" in pkg.last_change.parts_touched

    back = reloaded(pkg)
    defined = next(
        s for s in back.style_definitions_part.contents.style if s.style_id == "Hyperlink"
    )
    assert defined.based_on.val == "DefaultParagraphFont"
    assert str(defined.r_pr.color.theme_color.value) == "hyperlink"
    assert str(defined.r_pr.u.val.value) == "single"


def test_a_document_that_has_the_style_keeps_its_styles_part_byte_for_byte():
    source = FIXTURES / "hyperlink_dupe.docx"
    pkg = load(source)
    pkg.body.paragraphs[0].get_range().hyperlink = "https://example.com/"
    assert "/word/styles.xml" not in pkg.last_change.parts_touched
    with zipfile.ZipFile(source) as archive:
        original = archive.read("word/styles.xml")
    assert part_bytes(pkg.save())["word/styles.xml"] == original


# ---------------------------------------------------------------------------
# docx4j-core-ts check 16: a link over part of an existing link
# ---------------------------------------------------------------------------

WORD_CHECK_16_LINKS = {
    "F": ["https://example.com/f2"] * 3,
    "G": ["", "https://example.com/g2", ""],
    "H": ["", "", ""],
}
WORD_CHECK_16_XML = {
    "F": '<w:p><w:r><w:t xml:space="preserve">Case F: </w:t></w:r><w:hyperlink r:id="" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>alpha beta gamma</w:t></w:r></w:hyperlink></w:p>',
    "G": '<w:p><w:r><w:t xml:space="preserve">Case G: </w:t></w:r><w:r><w:t xml:space="preserve">alpha </w:t></w:r><w:hyperlink r:id="" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>beta</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> gamma</w:t></w:r></w:p>',
    "H": '<w:p><w:r><w:t xml:space="preserve">Case H: </w:t></w:r><w:r><w:t>alpha beta gamma</w:t></w:r></w:p>',
}


def test_a_link_over_part_of_an_existing_link_does_what_word_does_check_16():
    pkg = create_package()
    steps = {
        "F": [("alpha beta", "https://example.com/f1"), ("beta gamma", "https://example.com/f2")],
        "G": [("alpha beta gamma", "https://example.com/g1"), ("beta", "https://example.com/g2")],
        "H": [("alpha beta gamma", "https://example.com/h1"), ("beta", "")],
    }
    paragraphs = {}
    for case, actions in steps.items():
        paragraph = pkg.body.insert_paragraph(f"Case {case}: alpha beta gamma")
        paragraphs[case] = paragraph
        for target, address in actions:
            paragraph.search(target, match_case=True)[0].hyperlink = address
        reads = [paragraph.search(word)[0].hyperlink for word in ("alpha", "beta", "gamma")]
        assert reads == WORD_CHECK_16_LINKS[case], case
        assert_parents(paragraph)
        assert comparable(paragraph.get_xml()) == comparable(WORD_CHECK_16_XML[case]), case

    # F: one link over the union, and "Case F: " untouched
    f = paragraphs["F"]
    links = [
        item
        for item in f.element.content
        if item is not None and type(item).__name__ == "PHyperlink"
    ]
    assert len(links) == 1
    assert f.element.content[0].r_pr is None


# ---------------------------------------------------------------------------
# docx4j-core-ts check 19: a link set at a caret
# ---------------------------------------------------------------------------


def test_a_caret_is_refused_as_office_js_refuses_it_and_reads_as_it_reads_check_19():
    pkg = create_package()

    def set_up(case, link):
        paragraph = pkg.body.insert_paragraph(f"Case {case}: alpha beta gamma")
        if link:
            paragraph.search(link[0], match_case=True)[0].hyperlink = link[1]
        return paragraph

    def in_beta(paragraph):
        beta = paragraph.search("beta")[0]
        return Range(paragraph, beta.start + 1, beta.start + 1)

    def after_alpha(paragraph):
        return paragraph.search("alpha")[0].get_range("End")

    cases = [
        (
            "I",
            ("alpha beta gamma", "https://example.com/i1"),
            in_beta,
            "https://example.com/i2",
            "https://example.com/i1",
        ),
        ("J", None, in_beta, "https://example.com/j1", ""),
        (
            "K",
            ("alpha", "https://example.com/k1"),
            after_alpha,
            "https://example.com/k2",
            "https://example.com/k1",
        ),
        (
            "L",
            ("alpha beta gamma", "https://example.com/l1"),
            in_beta,
            "",
            "https://example.com/l1",
        ),
    ]
    rels = pkg.main_document_part.relationships_part
    for case, link, caret_of, address, reads in cases:
        paragraph = set_up(case, link)
        caret = caret_of(paragraph)
        assert caret.hyperlink == reads, case
        before = paragraph.get_xml()
        count = len(rels)
        with pytest.raises(SpanError) as raised:
            caret.hyperlink = address
        assert raised.value.code == "range.empty", case
        assert paragraph.get_xml() == before, f"case {case}: nothing changed"
        assert len(rels) == count, f"case {case}: no relationship added"


def test_only_the_links_the_range_touches_are_removed():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("one two three")

    def word(text):
        return paragraph.search(text)[0]

    word("one").hyperlink = "https://example.com/one"
    word("three").hyperlink = "https://example.com/three"
    word("two").hyperlink = "https://example.com/two"
    assert [word(w).hyperlink for w in ("one", "two", "three")] == [
        "https://example.com/one",
        "https://example.com/two",
        "https://example.com/three",
    ]

    word("two").hyperlink = ""
    assert [word(w).hyperlink for w in ("one", "two", "three")] == [
        "https://example.com/one",
        "",
        "https://example.com/three",
    ]
    assert paragraph.text == "one two three"
    styled = [getattr(getattr(run.r_pr, "r_style", None), "val", None) for run in paragraph.runs]
    assert styled.count("Hyperlink") == 2, styled

    # an empty range reads the link it lies in, and nothing where there is none
    assert Range(paragraph, 9, 9).hyperlink == "https://example.com/three"
    assert Range(paragraph, 5, 5).hyperlink == ""


# ---------------------------------------------------------------------------
# the relationship belongs to the range's part
# ---------------------------------------------------------------------------


def test_a_header_paragraphs_link_is_related_from_the_header_part():
    pkg = fixture("header-no-rels.docx")  # its headers have no .rels of their own
    header = next(h for h in pkg.header_parts() if str(h.part_name) == "/word/header1.xml")
    assert header.relationships_part is None
    main_rels = len(pkg.main_document_part.relationships_part)

    span = header.body.paragraphs[0].search("header")[0]
    span.hyperlink = "https://example.com/header"
    assert span.hyperlink == "https://example.com/header"
    rels = header.relationships_part
    assert rels is not None, "the header part's .rels is created"
    [rel] = rels.get_relationships_by_type(HYPERLINK_REL)
    assert rel.target == "https://example.com/header"
    assert len(pkg.main_document_part.relationships_part) == main_rels, "nothing on the main part"
    assert "/word/_rels/header1.xml.rels" in pkg.last_change.parts_touched

    data = pkg.save()
    assert "https://example.com/header" in entry(data, "word/_rels/header1.xml.rels")
    back = load(BytesIO(data))
    again = next(h for h in back.header_parts() if str(h.part_name) == "/word/header1.xml")
    assert again.body.paragraphs[0].search("header")[0].hyperlink == "https://example.com/header"


def test_an_address_needs_a_part_and_a_location_does_not():
    from docx4j_py.model.content.body import Body
    from docx4j_py.wml import Body as BodyElement

    holder = BodyElement()
    detached = Body(None, holder)
    paragraph = detached.insert_paragraph("no part behind this")
    paragraph.get_range().hyperlink = "#here"
    assert paragraph.get_range().hyperlink == "#here"
    before = paragraph.get_xml()
    with pytest.raises(ContentError) as raised:
        paragraph.search("part")[0].hyperlink = "https://example.com/"
    assert raised.value.code == "hyperlink.no_part"
    assert paragraph.get_xml() == before, "refused before the link it touches was unwrapped"


# ---------------------------------------------------------------------------
# the refusals, made before anything changes
# ---------------------------------------------------------------------------


def _state(pkg, paragraph):
    return paragraph.get_xml(), len(pkg.main_document_part.relationships_part)


def test_a_span_crossing_a_tracked_insertion_is_refused():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("before ")
    pkg.author = Author("Ada")
    pkg.change_tracking_mode = "TrackAll"
    paragraph.insert_text("inserted", location="End")
    pkg.change_tracking_mode = "Off"
    assert "<w:ins " in paragraph.get_xml()

    before = _state(pkg, paragraph)
    with pytest.raises(SpanError) as raised:
        paragraph.search("before inserted")[0].hyperlink = "https://example.com/"
    assert raised.value.code == "range.crosses_holder"
    assert _state(pkg, paragraph) == before

    # wholly inside the insertion is fine; the insertion goes inside the link,
    # since a w:ins cannot hold a w:hyperlink (CT_RunTrackChange)
    paragraph.search("inserted")[0].hyperlink = "https://example.com/"
    assert_parents(paragraph)
    xml = paragraph.get_xml()
    assert re.search(r'<w:hyperlink [^>]*><w:ins w:id="1" ', xml), xml
    assert "<w:ins" not in xml.split("<w:hyperlink")[0], "nothing of it left outside"


def test_a_link_over_part_of_an_insertion_splits_the_insertion_around_it():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("x")
    pkg.author = Author("Ada")
    pkg.change_tracking_mode = "TrackAll"
    paragraph.insert_text(" one two three", location="End")
    pkg.change_tracking_mode = "Off"

    paragraph.search("two")[0].hyperlink = "https://example.com/"
    assert paragraph.text == "x one two three"
    assert paragraph.search("two")[0].hyperlink == "https://example.com/"
    assert paragraph.search("one")[0].hyperlink == ""
    xml = paragraph.get_xml()
    # the part before keeps the element and its id; the link's and the part
    # after are revisions of the same kind and author with ids of their own
    found = re.findall(r'<w:ins w:id="(\d+)" w:author="Ada"', xml)
    assert found[0] == "1" and len(found) == 3 and len(set(found)) == 3, xml
    assert re.search(r"<w:hyperlink [^>]*><w:ins [^>]*><w:r>.*?<w:t>two</w:t>", xml), xml
    assert_parents(paragraph)
    added = [c for c in pkg.body.get_tracked_changes() if c.type == "Added"]
    assert "".join(c.text for c in added) == " one two three"
    assert paragraph.search("three")[0].get_tracked_changes(), "the part after is still a revision"

    # rejecting it all leaves the link empty, as Word leaves that form
    pkg.body.reject_all()
    assert pkg.body.paragraphs[0].text == "x"


def test_a_span_around_a_tracked_deletion_is_refused():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("the very quick fox")
    pkg.author = Author("Ada")
    pkg.change_tracking_mode = "TrackAll"
    paragraph.search("very ")[0].delete()
    pkg.change_tracking_mode = "Off"
    assert paragraph.text == "the quick fox"

    before = _state(pkg, paragraph)
    with pytest.raises(SpanError) as raised:
        paragraph.search("the quick")[0].hyperlink = "https://example.com/"
    assert raised.value.code == "range.crosses_holder"
    assert "w:del" in raised.value.message
    assert _state(pkg, paragraph) == before


def test_a_span_crossing_a_content_control_is_refused_and_a_link_it_touches_survives():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("see this page now")
    paragraph.search("this page")[0].insert_content_control("RichText")
    paragraph.search("this")[0].hyperlink = "https://example.com/inside"

    before = _state(pkg, paragraph)
    with pytest.raises(SpanError) as raised:
        paragraph.search("see this")[0].hyperlink = "https://example.com/across"
    assert raised.value.code == "range.crosses_holder"
    # docx4j-core-ts unwraps the touched link before it refuses; here nothing moves
    assert _state(pkg, paragraph) == before
    assert paragraph.search("this")[0].hyperlink == "https://example.com/inside"


def test_a_bookmark_between_the_runs_goes_into_the_link_and_no_text_moves():
    """docx4j-core-ts replaces as many items as there are runs, from the first run on.

    With a ``w:bookmarkStart`` between two runs that drops the bookmark and
    leaves the second run twice ("alpha betabeta gamma"); here the slice from
    the first run to the last goes into the link, markers and all.
    """
    pkg = create_package()
    pkg.body.add_object(
        el.p(
            content=[
                r("alpha "),
                el.bookmarkStart(id=5, name="bm"),
                r("beta"),
                el.bookmarkEnd(id=5),
                r(" gamma"),
            ]
        )
    )
    paragraph = last(pkg)
    paragraph.search("alpha beta")[0].hyperlink = "https://example.com/"
    assert paragraph.text == "alpha beta gamma"
    assert_parents(paragraph)
    xml = paragraph.get_xml()
    assert re.search(
        r'<w:hyperlink [^>]*><w:r>.*<w:bookmarkStart w:id="5" w:name="bm"/>.*</w:hyperlink>', xml
    ), xml
    assert xml.count("<w:bookmarkStart") == 1 and xml.count("<w:bookmarkEnd") == 1
    assert paragraph.search("gamma")[0].hyperlink == ""


def test_a_span_crossing_a_links_edge_takes_the_link_in_rather_than_being_refused():
    pkg = fixture("hyperlink.docx")
    paragraph = pkg.body.paragraphs[0]
    assert paragraph.text == "This document contains a hyperlink"
    span = paragraph.search("contains a hyperlink")[0]
    span.hyperlink = "https://example.com/"
    links = [item for item in paragraph.element.content if type(item).__name__ == "PHyperlink"]
    assert len(links) == 1
    assert Range(paragraph, span.start, span.end).hyperlink == "https://example.com/"
    assert paragraph.search("hyperlink")[0].hyperlink == "https://example.com/"
    assert paragraph.search("This")[0].hyperlink == ""


# ---------------------------------------------------------------------------
# the ChangeReport, dry_run, change tracking, to_api_script
# ---------------------------------------------------------------------------


def test_the_change_report_names_the_paragraph_the_values_and_the_parts():
    pkg = sample("2010-sample1.docx")
    paragraph = pkg.body.paragraphs[0]
    paragraph.search("first")[0].hyperlink = "https://example.com/"
    report = pkg.last_change
    assert report.operation == "hyperlink"
    assert report.addresses == (paragraph.address,)
    assert report.text_before == "hyperlink="
    assert report.text_after == "hyperlink=https://example.com/"
    assert report.parts_touched == (
        "/word/document.xml",
        "/word/_rels/document.xml.rels",
        "/word/styles.xml",
    )

    paragraph.search("first")[0].hyperlink = ""
    report = pkg.last_change
    assert report.text_before == "hyperlink=https://example.com/"
    assert report.text_after == "hyperlink="
    assert report.parts_touched == ("/word/document.xml",), "a removal relates nothing"


def test_a_refused_call_records_no_report():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("text")
    count = len(pkg.changes)
    with pytest.raises(SpanError):
        paragraph.get_range("End").hyperlink = "https://example.com/"
    assert len(pkg.changes) == count


def test_a_dry_run_leaves_the_package_byte_for_byte_and_the_commit_gets_the_trials_id():
    pkg = sample("2010-sample1.docx")
    _ = pkg.body
    _ = pkg.style_definitions_part.contents  # read either way by the style check
    before = part_bytes(pkg.save(), relationships=True)
    rels = pkg.main_document_part.relationships_part
    count = len(rels)

    with pkg.dry_run() as trial:
        span = trial.body.paragraphs[0].search("first")[0]
        span.hyperlink = "https://example.com/"
        assert span.hyperlink == "https://example.com/"
        trial_id = re.search(r'r:id="([^"]+)"', span.paragraph.get_xml()).group(1)
        assert trial.last_change.parts_touched[1] == "/word/_rels/document.xml.rels"

    assert len(rels) == count
    assert pkg.body.paragraphs[0].get_range().hyperlink == ""
    assert part_bytes(pkg.save(), relationships=True) == before

    pkg.body.paragraphs[0].search("first")[0].hyperlink = "https://example.com/"
    assert re.search(r'r:id="([^"]+)"', pkg.body.paragraphs[0].get_xml()).group(1) == trial_id


def test_a_dry_run_in_a_header_detaches_the_relationships_part_it_created():
    pkg = fixture("header-no-rels.docx")
    header = next(h for h in pkg.header_parts() if str(h.part_name) == "/word/header1.xml")
    with pkg.dry_run() as trial:
        # trial.bodies(), not trial.header_parts()[i].body, which is the real header's
        trial_body = next(b for b in trial.bodies() if str(b.part.part_name) == "/word/header1.xml")
        trial_body.paragraphs[0].search("header")[0].hyperlink = "https://example.com/"
        assert header.relationships_part is not None, "a trial shares the part graph"
    assert header.relationships_part is None
    assert header.body.paragraphs[0].get_range().hyperlink == ""


def test_under_change_tracking_the_restyle_is_a_formatting_revision():
    pkg = create_package()
    pkg.body.insert_paragraph("a tracked link")
    pkg.author = Author("Ada")
    pkg.change_tracking_mode = "TrackAll"
    pkg.body.paragraphs[0].search("link")[0].hyperlink = "https://example.com/"
    xml = pkg.body.paragraphs[0].get_xml()
    assert re.search(r'<w:rStyle w:val="Hyperlink"/><w:rPrChange [^>]*w:author="Ada"', xml), xml
    assert "<w:ins " not in xml and "<w:del " not in xml, "the wrap itself is not a revision"
    assert any(change.type == "Formatted" for change in pkg.body.get_tracked_changes())

    # rejecting it takes the style off; the w:hyperlink stays, Word having no markup for it
    pkg.change_tracking_mode = "Off"
    pkg.body.reject_all()
    assert pkg.body.paragraphs[0].search("link")[0].hyperlink == "https://example.com/"
    assert 'w:val="Hyperlink"' not in pkg.body.paragraphs[0].get_xml()


def test_a_link_over_this_authors_own_insertion_records_no_formatting_revision():
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("before ")
    pkg.author = Author("Ada")
    pkg.change_tracking_mode = "TrackAll"
    paragraph.insert_text("new words", location="End").hyperlink = "https://example.com/"
    xml = paragraph.get_xml()
    assert "<w:rPrChange" not in xml, xml
    assert re.search(r"<w:hyperlink [^>]*><w:ins [^>]*><w:r>", xml), xml


def test_a_link_made_and_removed_under_tracking_keeps_its_recorded_original():
    """Not docx4j-core-ts: it settles a w:rPrChange whose properties are back (its check 24).

    This package's ``Font`` setters do not settle either (CR-003 section 21.3 item 9),
    so the record stays, recording no properties --- a formatting revision that
    changes nothing, which Word itself would have dropped.
    """
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("a tracked link")
    pkg.author = Author("Ada")
    pkg.change_tracking_mode = "TrackAll"
    paragraph.search("link")[0].hyperlink = "https://example.com/"
    paragraph.search("link")[0].hyperlink = ""
    xml = paragraph.get_xml()
    assert "<w:hyperlink" not in xml
    assert "<w:rPrChange" in xml


def test_to_api_script_writes_a_linked_paragraph_as_one_insert_xml():
    """As docx4j-core-ts's toApiScript does: the setter is not emitted (CR-003 section 21.8)."""
    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("see the docs here")
    paragraph.search("the docs")[0].hyperlink = "https://docx4j.org/"
    script = paragraph.to_api_script()
    assert script.startswith("# w:hyperlink in the paragraph: as XML")
    assert ".hyperlink" not in script
