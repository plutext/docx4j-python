"""CR-037: an unknown ``w:`` element, read as Word 365 reads it (core-ts's CR-026 D7 and D8).

    .venv-fork/bin/python -m pytest tests/openpackaging/test_read_through.py -q

The oracle is core-ts's Word check 76 (``tests/fixtures/check76/``): 72 sources, each
beside Word 365's save but for the ten Word refuses. Rounds 6 to 8 are bridge saves,
so only their stories are compared, and these tests compare no more than that. A save
changes here, so the save's own Word check is in ``tests/README.md``.
"""

from __future__ import annotations

import hashlib
import io
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon

from docx4j_py import create_package, load
from docx4j_py.child import iter_children
from docx4j_py.el_index import WML_ELEMENTS
from docx4j_py.fragments import FragmentError, wml
from docx4j_py.model.text import TextOptions, extract, plain_text
from docx4j_py.openpackaging import LoadOptions, breaks, read_through
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.parts.wml import (
    CommentsPart,
    EndnotesPart,
    FooterPart,
    FootnotesPart,
    HeaderPart,
)
from docx4j_py.traversal import text_of

CHECK76 = ROOT / "tests" / "fixtures" / "check76"
SAMPLES = ROOT / "samples"
W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"
MADE_UP = W + "madeUp"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
DOCUMENT = "/word/document.xml"
NS = (
    f'xmlns:w="{W_URI}" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:ext="urn:ext"'
)
STORIES = TextOptions(include=["headers", "footers"], comments=True, notes=True)

#: Each file read as Word reads it: its part, and that part's report, each entry a name and what held it.
def through(holder: str, part: str = DOCUMENT) -> tuple[str, list[tuple[str, str]]]:
    return part, [(MADE_UP, W + holder)]


def twice(holder: str) -> tuple[str, list[tuple[str, str]]]:
    return DOCUMENT, [(MADE_UP, W + holder), (MADE_UP, MADE_UP)]


AS_WORD = {
    "76t-madeup-holding-a-run": through("p"),
    "76t1-madeup-holding-p-in-body": through("body"),
    "76t2-madeup-holding-p-in-tc": through("tc"),
    "76t3-madeup-holding-t-in-r": through("r"),
    "76t4-madeup-holding-tr-in-tbl": through("tbl"),
    "76t7-madeup-in-madeup": twice("p"),
    "76t8-madeup-with-no-ignorable": through("p"),
    "76t9-madeup-holding-tc-in-tr": through("tr"),
    "76t10-madeup-in-block-control": through("sdtContent"),
    "76t11-madeup-in-run-control": through("sdtContent"),
    "76t12-madeup-in-row-control": through("sdtContent"),
    "76t13-madeup-in-cell-control": through("sdtContent"),
    "76t14-madeup-in-hyperlink": through("hyperlink"),
    "76t15-madeup-in-ins": through("ins"),
    "76t16-madeup-in-hdr": through("hdr", "/word/header1.xml"),
    "76t17-madeup-in-comment": through("comment", "/word/comments.xml"),
    "76t19-madeup-in-fldSimple": through("fldSimple"),
    "76t20-madeup-in-smartTag": through("smartTag"),
    "76t21-madeup-in-block-customXml": through("customXml"),
    "76t22-madeup-in-run-customXml": through("customXml"),
    "76t23-madeup-in-row-customXml": through("customXml"),
    "76t24-madeup-in-cell-customXml": through("customXml"),
    "76t25-madeup-in-footnote": through("footnote", "/word/footnotes.xml"),
    "76t27-madeup-holding-a-run-in-pPr": through("pPr"),
    "76t29-madeup-in-ftr": through("ftr", "/word/footer1.xml"),
    "76t30-madeup-in-endnote": through("endnote", "/word/endnotes.xml"),
    "76t31-madeup-in-del": through("del"),
    "76t32-madeup-in-moveFrom": through("moveFrom"),
    "76t33-madeup-in-moveTo": through("moveTo"),
    "76t35-madeup-in-madeup-in-tc": twice("tc"),
    "76t36-madeup-in-madeup-in-hyperlink": twice("hyperlink"),
    "76t37-two-runs-in-pPr": through("pPr"),
    "76t38-run-and-property-in-pPr": (DOCUMENT, [(MADE_UP, W + "CT_PPr")]),  # the element stands, so the entry is the fork's
    "76t39-run-in-pPr-no-other-run": through("pPr"),
    "76t40-run-in-pPrChange": through("pPr"),
    "76t41-text-in-madeup": through("p"),
    "76t42-break-in-madeup": through("p"),
    "76t43-paragraph-in-madeup-in-p": (DOCUMENT, [(MADE_UP, W + "p"), (W + "p", MADE_UP)]),
    "76t56-two-held-paragraphs": (DOCUMENT, [(MADE_UP, W + "p"), (W + "p", MADE_UP)]),
    "76t58-held-paragraph-centred": (DOCUMENT, [(MADE_UP, W + "p"), (W + "p", MADE_UP)]),
    "76t68-host-right-held-centred": (DOCUMENT, [(MADE_UP, W + "p"), (W + "p", MADE_UP)]),
    "76t69-host-right-held-plain": (DOCUMENT, [(MADE_UP, W + "p"), (W + "p", MADE_UP)]),
    "76t70-held-paragraph-two-deep": (DOCUMENT, [(MADE_UP, W + "p"), (MADE_UP, MADE_UP), (W + "p", MADE_UP)]),
    "76t61-text-in-madeup-in-body": through("body"),
    "76t62-text-in-madeup-in-tc": through("tc"),
    "76t63-text-in-madeup-in-r": through("r"),
    "76t64-text-in-madeup-in-hyperlink": through("hyperlink"),
    "76t65-text-in-madeup-in-ins": through("ins"),
    "76t66-text-in-madeup-in-block-control": through("sdtContent"),
    "76t67-text-and-a-run-in-madeup": through("p"),
    "76t44-madeup-in-choice": (DOCUMENT, [(MADE_UP, MC + "Choice")]),
    "76t45-madeup-in-fallback": (DOCUMENT, [(MADE_UP, MC + "Fallback")]),
    "76t50-run-in-sectPr": (DOCUMENT, [(MADE_UP, W + "sectPr"), (W + "sectPr", W + "pPr")]),
    "76t51-second-name-in-p": (DOCUMENT, [(W + "neverDeclared", W + "p")]),
}
#: Where the words Word shows are ones a read here didn't have before.
BREAK = "76t42-break-in-madeup"
NO_WORDS = {"76t41-text-in-madeup", BREAK}
#: Where a save still loses what the element held beside its run, as it does a property's.
KEEPS_A_DROP = {"76t38-run-and-property-in-pPr"}
#: Where nothing moves: each file, and the fork's report for it.
UNMOVED = {
    "76t5-madeup-holding-jc-in-pPr": [("element", MADE_UP, W + "CT_PPr")],
    "76t6-madeup-holding-b-in-rPr": [("element", MADE_UP, W + "CT_RPr")],
    "76t18-hosts-control": [],
    "76t26-more-hosts-control": [],
    "76t34-last-hosts-control": [],
    "76t52-round-8-control": [],
    "76t60-round-9-control": [],
}
#: Files Word refuses, so with no save: each, and what the fork's report calls the holder.
REFUSED = {
    "76t28-madeup-holding-a-run-in-rPr": "CT_RPr",
    "76t46-run-in-sdtPr": "CT_SdtPr",
    "76t47-run-in-trPr": "CT_TrPr",
    "76t48-run-in-tcPr": "CT_TcPr",
    "76t49-run-in-tblPr": "CT_TblPr",
}
#: Files Word refuses with a w:p held in a w:p: REQ-316's shape it doesn't read, so the second parse's to skip and list.
HELD_REFUSED = {
    "76t53-held-paragraph-then-a-run": 1,
    "76t54-held-paragraph-first": 1,
    "76t55-held-paragraph-alone": 1,
    "76t57-two-held-paragraphs-then-a-run": 2,
    "76t59-both-aligned-then-a-run": 1,
}
#: Beyond the WD: REQ-316 reaches a w:p, and Word reads this in a w:hyperlink too. Pinned as it reads here, the words lost.
IN_A_HYPERLINK = "76t71-held-paragraph-in-hyperlink"
RUN_IN_PPR = "76t27-madeup-holding-a-run-in-pPr"

STRUCTURE = {
    W + n
    for n in ("pPr", "rPr", "tblPr", "trPr", "tcPr", "sdtPr", "sdtEndPr", "sectPr", "tblGrid", "tblPrEx", "smartTagPr", "customXmlPr")
}

TRACKED = {W + "ins": "ins", W + "del": "del", W + "moveFrom": "from", W + "moveTo": "to"}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def walk(obj: Any, out: list[str], mark: str = "t") -> None:
    """Text merged as Word merges runs, a tracked run's marked as such, with rows, cells and paragraph ends.

    Field characters and instructions aren't tokens: Word saves a ``w:fldSimple`` as
    a complex field, with or without the element (``76t26``).
    """
    for qname, child in iter_children(obj):
        if qname in STRUCTURE:
            continue
        if qname == W + "r":
            for name, item in iter_children(child):
                if name in (W + "t", W + "delText"):
                    head = mark + ":"
                    if out and out[-1].startswith(head):
                        out[-1] += item.value or ""
                    else:
                        out.append(head + (item.value or ""))
                elif name == W + "br":
                    out.append("br")
            continue
        if qname in (W + "tr", W + "tc"):
            out.append(qname[len(W) :])
        walk(child, out, TRACKED.get(qname, mark))
        if qname == W + "p":
            out.append("¶")


def tokens(pkg: Any) -> dict[str, list[str]]:
    """Each story's tokens: the body, the headers, the footers, the comments and the real notes."""
    stories: dict[str, list[str]] = {"body": []}
    walk(pkg.main_document_part.contents.body, stories["body"])
    for part in pkg.parts.parts():
        if isinstance(part, HeaderPart):
            walk(part.contents, stories.setdefault("headers", []))
        elif isinstance(part, FooterPart):
            walk(part.contents, stories.setdefault("footers", []))
        elif isinstance(part, CommentsPart):
            walk(part.contents, stories.setdefault("comments", []))
        elif isinstance(part, FootnotesPart):
            for note in part.contents.footnote:
                if note.type_value is None:
                    walk(note, stories.setdefault("notes", []))
        elif isinstance(part, EndnotesPart):
            for note in part.contents.endnote:
                if note.type_value is None:
                    walk(note, stories.setdefault("endnotes", []))
    return stories


def text(pkg: Any) -> str:
    return plain_text(extract(pkg, STORIES))


def source(stem: str) -> Path:
    return CHECK76 / f"{stem}.docx"


def word(stem: str) -> Any:
    return load(CHECK76 / f"{stem}-word365.docx")


def part_of(pkg: Any, name: str) -> Any:
    return next(part for part in (pkg.main_document_part, *pkg.parts.parts()) if str(part.part_name) == name)


def entries(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def saved(pkg: Any, **kwargs: Any) -> bytes:
    out = io.BytesIO()
    pkg.save(out, **kwargs)
    return out.getvalue()


def package_of(body_xml: str, **options: Any) -> Any:
    pkg = create_package()
    if options:
        pkg.load_options = LoadOptions(**options)
    pkg.main_document_part.set_bytes(
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'.encode()
    )
    return pkg


def body_tokens(body_xml: str, **options: Any) -> tuple[list[str], Any]:
    pkg = package_of(body_xml, **options)
    out: list[str] = []
    walk(pkg.main_document_part.contents.body, out)
    return out, pkg


def entry(item: Any) -> tuple[str, str]:
    """A report's name and holder. The fork's type name takes whichever namespace its metadata met first, so it's read by its local name."""
    local = item.parent_qname.rsplit("}", 1)[-1]
    return item.qname, W + local if local.startswith("CT_") else item.parent_qname


def names(pkg: Any) -> list[tuple[str, str]]:
    """The body's ``part.skipped``, each as its name and its holder's."""
    return [(item.qname[len(W) :], item.parent_qname.rsplit("}", 1)[-1]) for item in pkg.main_document_part.skipped]


def tree(xml: str) -> etree._Element:
    """`xml` as a tree, the namespaces declared on its first element."""
    at = xml.index(">") if xml[xml.index(">") - 1] != "/" else xml.index(">") - 1
    return etree.fromstring(f"{xml[:at]} {NS}{xml[at:]}".encode())


def inner(root: etree._Element) -> str:
    """`root`'s content with the test's namespace declarations left out."""
    out = (root.text or "") + "".join(etree.tostring(child, encoding="unicode") for child in root)
    for declaration in NS.split(" "):
        out = out.replace(f" {declaration}", "")
    return out


R = '<w:r><w:t xml:space="preserve">{}</w:t></w:r>'
A, B, C = R.format("a"), R.format("b"), R.format("c")


# ---------------------------------------------------------------------------
# the fixtures
# ---------------------------------------------------------------------------


def test_the_files_are_check_76s() -> None:
    listed = dict(reversed(line.split()) for line in (CHECK76 / "SHA1SUMS").read_text().splitlines())
    assert set(listed) == {path.name for path in CHECK76.glob("*.docx")}
    assert len(listed) == 144
    for name, digest in listed.items():
        assert hashlib.sha1((CHECK76 / name).read_bytes()).hexdigest() == digest, name
    sources = [path.stem for path in CHECK76.glob("76t*.docx") if not path.stem.endswith("-word365")]
    assert set(sources) == {*AS_WORD, *UNMOVED, *REFUSED, *HELD_REFUSED, IN_A_HYPERLINK}
    assert len(sources) == 72
    assert not any((CHECK76 / f"{stem}-word365.docx").exists() for stem in (*REFUSED, *HELD_REFUSED))


@pytest.mark.parametrize("stem", AS_WORD)
def test_each_file_reads_as_words_save(stem: str) -> None:
    ours, saves = load(source(stem)), word(stem)
    assert tokens(ours) == tokens(saves)
    assert text(ours) == text(saves)
    assert saves.skipped == [] and saves.warnings == []


@pytest.mark.parametrize("stem", AS_WORD)
def test_each_file_says_what_was_read_through(stem: str) -> None:
    name, expected = AS_WORD[stem]
    pkg = load(source(stem))
    part = part_of(pkg, name)
    assert part.contents is not None
    assert [entry(item) for item in part.skipped] == expected
    assert all(item.kind == "element" for item in part.skipped)
    assert [(w.breaks, w.text_elements) for w in part.warnings] == ([(1, 0)] if stem == BREAK else [])
    text(pkg)  # every story read
    assert [(str(where), item) for where, item in pkg.skipped] == [(name, item) for item in part.skipped]


@pytest.mark.parametrize("stem", UNMOVED)
def test_where_nothing_moves_the_file_reads_as_before(stem: str) -> None:
    ours, saves = load(source(stem)), word(stem)
    assert tokens(ours) == tokens(saves)
    assert text(ours) == text(saves)
    assert [(item.kind, *entry(item)) for item in ours.main_document_part.skipped] == UNMOVED[stem]


def test_a_run_in_a_w_ppr_goes_to_the_head_of_the_paragraph() -> None:
    """``76t27``: ``HIDDEN WORDSTARGET TEXT``, one run in Word's save."""
    pkg = load(source(RUN_IN_PPR))
    assert pkg.body.paragraphs[1].text == "HIDDEN WORDSTARGET TEXT"
    assert word(RUN_IN_PPR).body.paragraphs[1].text == "HIDDEN WORDSTARGET TEXT"


@pytest.mark.parametrize(("stem", "holder"), REFUSED.items())
def test_a_file_word_refuses_stays_dropped_and_reported(stem: str, holder: str) -> None:
    """``76t28`` and ``76t46`` to ``76t49``: read here, as any is, the words gone, the element listed."""
    pkg = load(source(stem))
    assert not any("HIDDEN WORDS" in token for token in tokens(pkg)["body"][2:])  # the first paragraph quotes them
    assert [entry(item) for item in pkg.main_document_part.skipped] == [(MADE_UP, W + holder)]


# ---------------------------------------------------------------------------
# the save and the round trip
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("stem", AS_WORD)
def test_a_save_writes_what_word_reads(stem: str) -> None:
    pkg = load(source(stem))
    assert part_of(pkg, AS_WORD[stem][0]).contents is not None  # the text model reads privately, marking nothing
    data = saved(pkg)
    name = AS_WORD[stem][0].lstrip("/")
    assert b"<w:madeUp" not in entries(data)[name] and b"<w:neverDeclared" not in entries(data)[name]
    again = load(data)
    assert tokens(again) == tokens(word(stem))
    assert text(again) == text(word(stem))
    assert again.skipped == [] and again.warnings == []


@pytest.mark.parametrize("normalise", [True, False])
@pytest.mark.parametrize("stem", [stem for stem in AS_WORD if stem not in KEEPS_A_DROP])
def test_a_save_is_its_repaired_source(stem: str, normalise: bool) -> None:
    """Section 3.5: the save holds nothing but the repair, the repair counted ``read-through``."""
    name = AS_WORD[stem][0].lstrip("/")
    pkg = load(source(stem))
    assert part_of(pkg, AS_WORD[stem][0]).contents is not None
    new = entries(saved(pkg, normalise=normalise))[name]
    original = entries(source(stem).read_bytes())[name]
    report = canon.compare(original, new, normalise=normalise, repaired=True)
    assert report.identical, report.differences
    assert report.counts.get("read-through") == sum(1 for entry, _holder in AS_WORD[stem][1] if entry in (MADE_UP, W + "neverDeclared"))
    assert not canon.compare(original, new, normalise=normalise).identical


@pytest.mark.parametrize("stem", ["76t-madeup-holding-a-run", RUN_IN_PPR, "76t16-madeup-in-hdr"])
def test_a_part_never_read_saves_byte_for_byte(stem: str) -> None:
    original = entries(source(stem).read_bytes())
    pkg = load(source(stem))
    assert "HIDDEN WORDS" in text(pkg)  # the text model's read is private
    data = entries(saved(pkg))
    for name in ("word/document.xml", "word/header1.xml"):
        if name in original:
            assert data[name] == original[name]


# ---------------------------------------------------------------------------
# what it takes, and what it leaves
# ---------------------------------------------------------------------------


def test_the_names_are_the_schemas_elements() -> None:
    """The table is the model's, and the schema is the check: a real name read through would lose its element."""
    schema = etree.parse(ROOT / "schemas" / "wml" / "wml.xsd").getroot()
    declared = {
        element.get("name")
        for element in schema.iter("{http://www.w3.org/2001/XMLSchema}element")
        if element.get("name")
    }
    assert WML_ELEMENTS == declared
    assert "uniqueTag" in WML_ELEMENTS  # a simple type, so no class
    assert "madeUp" not in WML_ELEMENTS


def test_the_hosts_are_the_twenty_measured() -> None:
    measured = (
        "p", "body", "tc", "r", "tbl", "tr", "sdtContent", "hyperlink", "ins", "del", "moveFrom", "moveTo", "hdr", "ftr",
        "comment", "fldSimple", "smartTag", "customXml", "footnote", "endnote",
    )  # fmt: skip
    assert read_through._HOSTS == {W + local for local in measured}


@pytest.mark.parametrize(
    ("kind", "qname", "fires"),
    [
        ("element", MADE_UP, True),
        ("attribute", MADE_UP, False),
        ("element", W + "jc", False),
        ("element", "{urn:ext}opaque", False),
        ("element", "{http://purl.oclc.org/ooxml/wordprocessingml/main}madeUp", False),
        ("element", "madeUp", False),
        ("element", None, False),
    ],
)
def test_what_triggers_the_pass(kind: str, qname: str | None, fires: bool) -> None:
    class Item:
        pass

    item = Item()
    item.kind, item.qname = kind, qname
    assert read_through.triggers([item]) is fires


def test_an_empty_report_triggers_nothing() -> None:
    assert read_through.triggers([]) is False


def test_an_empty_element_just_goes() -> None:
    got, pkg = body_tokens(f'<w:p>{A}<w:madeUp w:val="1"/>{B}</w:p>')
    assert got == ["t:ab", "¶"]
    assert names(pkg) == [("madeUp", "p")]


def test_order_is_kept_and_text_held_directly_goes() -> None:
    """REQ-314. The host's own text is its, and goes with the parse as ever."""
    root = tree("<w:p>one<w:bookmarkStart/>two<w:madeUp>three<!--c--><w:r/>four<w:r/>five</w:madeUp>six<w:bookmarkEnd/>seven</w:p>")
    assert read_through.repair(root) == 1
    assert inner(root) == "one<w:bookmarkStart/>two<!--c--><w:r/><w:r/>six<w:bookmarkEnd/>seven"


def test_the_tail_after_an_element_is_kept_with_nothing_before_it() -> None:
    root = tree("<w:p><w:madeUp>one</w:madeUp>two</w:p>")
    assert read_through.repair(root) == 1
    assert inner(root) == "two"


def test_the_path_is_as_read_for_nested_elements() -> None:
    pkg = load(source("76t7-madeup-in-madeup"))
    assert pkg.main_document_part.contents is not None
    outer, nested = pkg.main_document_part.skipped
    assert outer.path == "/".join(W + step for step in ("document", "body", "p", "madeUp"))
    assert nested.path == "/".join(W + step for step in ("document", "body", "p", "madeUp", "madeUp"))


def test_a_known_element_in_the_wrong_place_is_left() -> None:
    got, pkg = body_tokens(f"<w:p><w:jc w:val='center'>{A}</w:jc>{B}</w:p>")
    assert got == ["t:b", "¶"]
    assert names(pkg) == [("jc", "p")]


def test_a_name_the_standard_defines_and_the_model_lacks_goes_as_it_did() -> None:
    """``w:contentPart`` is the one such name a host can hold, and it holds nothing, so the drop and the pass agree."""
    assert "contentPart" not in WML_ELEMENTS
    got, pkg = body_tokens('<w:p><w:r><w:contentPart xmlns:r="urn:r" r:id="rId1"/><w:t>a</w:t></w:r></w:p>')
    assert got == ["t:a", "¶"]
    assert names(pkg) == [("contentPart", "r")]


def test_a_madeup_in_a_known_element_in_the_wrong_place_is_not_taken() -> None:
    root = tree(f"<w:p><w:jc><w:madeUp>{A}</w:madeUp></w:jc></w:p>")
    assert read_through.repair(root) == 0


@pytest.mark.parametrize(
    "xml",
    [
        f"<w:sdt><w:sdtPr><w:madeUp>{A}</w:madeUp></w:sdtPr></w:sdt>",
        f"<w:tr><w:trPr><w:madeUp>{A}</w:madeUp></w:trPr></w:tr>",
        f"<w:tc><w:tcPr><w:madeUp>{A}</w:madeUp></w:tcPr></w:tc>",
        f"<w:tbl><w:tblPr><w:madeUp>{A}</w:madeUp></w:tblPr></w:tbl>",
        f"<w:body><w:sectPr><w:madeUp>{A}</w:madeUp></w:sectPr></w:body>",
        f"<w:docPartBody><w:madeUp><w:p>{A}</w:p></w:madeUp></w:docPartBody>",
        f"<w:txbxContent><w:madeUp><w:p>{A}</w:p></w:madeUp></w:txbxContent>",
        f"<w:document><w:madeUp>{A}</w:madeUp></w:document>",
        f"<w:settings><w:madeUp>{A}</w:madeUp></w:settings>",
    ],
)
def test_a_host_not_measured_is_left(xml: str) -> None:
    root = tree(xml)
    before = etree.tostring(root)
    assert read_through.repair(root) == 0
    assert etree.tostring(root) == before


@pytest.mark.parametrize(
    "xml",
    [
        "<w:p><w:pPr><w:rPr><w:ins><w:madeUp/></w:ins></w:rPr></w:pPr></w:p>",
        "<w:r><w:rPr><w:ins><w:madeUp/></w:ins></w:rPr></w:r>",
        "<w:tr><w:trPr><w:ins><w:madeUp/></w:ins></w:trPr></w:tr>",
        "<w:p><w:pPr><w:numPr><w:ins><w:madeUp/></w:ins></w:numPr></w:pPr></w:p>",
        "<m:ctrlPr><w:ins><w:madeUp><w:rPr/></w:madeUp></w:ins></m:ctrlPr>",
        "<m:r><w:ins><w:madeUp><m:t>a</m:t></w:madeUp></w:ins></m:r>",
        '<w:settings><w:footnotePr><w:footnote w:id="0"><w:madeUp/></w:footnote></w:footnotePr></w:settings>',
        '<w:settings><w:endnotePr><w:endnote w:id="0"><w:madeUp/></w:endnote></w:endnotePr></w:settings>',
        "<w:r><w:rPr><w:del><w:madeUp/></w:del></w:rPr></w:r>",
        "<w:tr><w:trPr><w:del><w:madeUp/></w:del></w:trPr></w:tr>",
        "<m:ctrlPr><w:del><w:madeUp/></w:del></m:ctrlPr>",
        "<m:r><w:del><w:madeUp/></w:del></m:r>",
        "<w:r><w:rPr><w:moveFrom><w:madeUp/></w:moveFrom></w:rPr></w:r>",
        "<w:r><w:rPr><w:moveTo><w:madeUp/></w:moveTo></w:rPr></w:r>",
    ],
)
def test_a_hosts_name_standing_for_another_element_is_no_host(xml: str) -> None:
    """Six of the hosts are other elements under these parents, which no file measures."""
    root = tree(xml)
    assert read_through.repair(root) == 0


def test_the_other_elements_parents_are_the_schemas() -> None:
    """Every declaration of the two names but the host's, followed through its groups and types to what holds it.

    A declaration's own group isn't its parent: ``EG_RPrMath``, which reads like run
    properties, stands directly in ``m:ctrlPr``. The first list missed that.
    """
    xs = "{http://www.w3.org/2001/XMLSchema}"
    schemas = [etree.parse(ROOT / "schemas" / path).getroot() for path in ("wml/wml.xsd", "shared/shared-math-2ed.xsd")]

    def local(name: str | None) -> str | None:
        return name.rsplit(":", 1)[-1] if name else name

    def named_owner(node: etree._Element) -> etree._Element:
        return next(a for a in node.iterancestors() if a.get("name") and a.tag in (xs + "complexType", xs + "group", xs + "element"))

    def holders(owner: etree._Element, seen: set[str]) -> set[str]:
        """The elements `owner`, a named group, type or element, stands for."""
        name = owner.get("name")
        if owner.tag == xs + "element":
            return {f"{{{owner.getroottree().getroot().get('targetNamespace')}}}{name}"}
        if name in seen:
            return set()
        seen.add(name)
        found: set[str] = set()
        for schema in schemas:
            for node in schema.iter(xs + "element", xs + "group", xs + "extension"):
                if node.tag == xs + "element" and local(node.get("type")) == name and node.get("name"):
                    found |= holders(node, seen)
                elif local(node.get("ref") if node.tag == xs + "group" else node.get("base")) == name:
                    found |= holders(named_owner(node), seen)
        return found

    hosts = {"ins": None, "del": None, "moveFrom": None, "moveTo": None, "footnote": "CT_FtnEdn", "endnote": "CT_FtnEdn"}
    assert set(read_through._ANOTHER_ELEMENT_UNDER) == {W + local for local in hosts}
    for tag, host_type in ((W + local, kind) for local, kind in hosts.items()):
        derived: set[str] = set()
        for schema in schemas:
            for node in schema.iter(xs + "element"):
                other = node.get("type") not in (host_type, "CT_RunTrackChange") and node.getparent().tag != xs + "schema"
                if node.get("name") == tag[len(W) :] and other:
                    derived |= holders(named_owner(node), set())
        assert derived == read_through._ANOTHER_ELEMENT_UNDER[tag]


@pytest.mark.parametrize(
    ("xml", "kept"),
    [
        (f"<w:body><w:ins><w:madeUp>{A}</w:madeUp></w:ins></w:body>", "<w:ins>"),
        (f"<w:footnotes><w:footnote><w:madeUp><w:p>{A}</w:p></w:madeUp></w:footnote></w:footnotes>", "<w:footnote>"),
        (f"<w:endnotes><w:endnote><w:madeUp><w:p>{A}</w:p></w:madeUp></w:endnote></w:endnotes>", "<w:endnote>"),
        (f"<w:p><w:del><w:madeUp>{A}</w:madeUp></w:del></w:p>", "<w:del>"),
        (f"<w:p><w:moveFrom><w:madeUp>{A}</w:madeUp></w:moveFrom></w:p>", "<w:moveFrom>"),
        (f"<w:p><w:moveTo><w:madeUp>{A}</w:madeUp></w:moveTo></w:p>", "<w:moveTo>"),
    ],
)
def test_the_same_names_holding_content_are_hosts(xml: str, kept: str) -> None:
    root = tree(xml)
    assert read_through.repair(root) == 1
    assert kept in inner(root) and "madeUp" not in inner(root)


def test_an_element_of_another_namespace_stops_the_walk() -> None:
    """Word drops one with all it holds (``76k``), so what it holds isn't read through."""
    root = tree(f"<w:p><ext:opaque><w:madeUp>{A}</w:madeUp></ext:opaque></w:p>")
    assert read_through.repair(root) == 0


def test_an_element_of_another_namespace_held_by_one_taken_stays_the_parses() -> None:
    got, pkg = body_tokens(f"<w:p>{A}<w:madeUp><ext:opaque>{B}</ext:opaque>{C}</w:madeUp></w:p>")
    assert got == ["t:ac", "¶"]
    assert [item.qname for item in pkg.main_document_part.skipped] == ["{urn:ext}opaque", MADE_UP]


def test_a_strict_documents_namespace_is_not_matched() -> None:
    strict = "http://purl.oclc.org/ooxml/wordprocessingml/main"
    root = etree.fromstring(f'<w:p xmlns:w="{strict}"><w:madeUp><w:r/></w:madeUp></w:p>')
    assert read_through.repair(root) == 0


def test_what_its_host_doesnt_admit_is_the_second_parses() -> None:
    """Beside REQ-316 the pass doesn't look. A ``w:tc`` left in a ``w:p`` is skipped and listed."""
    got, pkg = body_tokens(f"<w:p>{A}<w:madeUp><w:tc><w:p>{B}</w:p></w:tc></w:madeUp>{C}</w:p>")
    assert got == ["t:ac", "¶"]
    assert [item.qname for item in pkg.main_document_part.skipped] == [W + "tc", MADE_UP]


# ---------------------------------------------------------------------------
# a w:p in such an element in a w:p (REQ-316, core-ts's D11 and round 9)
# ---------------------------------------------------------------------------


def test_a_held_paragraph_ends_the_host_and_an_empty_one_follows() -> None:
    got, pkg = body_tokens(f"<w:p>{A}<w:madeUp><w:p>{B}</w:p></w:madeUp></w:p><w:p>{C}</w:p>")
    assert got == ["t:ab", "¶", "¶", "t:c", "¶"]
    assert names(pkg) == [("madeUp", "p"), ("p", "madeUp")]


def test_a_further_held_paragraph_stands_alone() -> None:
    """``76t56``."""
    root = tree(f"<w:body><w:p>{A}<w:madeUp><w:p>{B}</w:p><w:p>{C}</w:p></w:madeUp></w:p></w:body>")
    dropped: list[Any] = []
    assert read_through.repair(root, dropped) == 1
    assert inner(root) == f"<w:p>{A}{B}</w:p><w:p>{C}</w:p><w:p/>"
    assert [item.qname for item in dropped] == [MADE_UP, W + "p"]  # a further one keeps its tags


@pytest.mark.parametrize(
    ("stem", "joined", "left"),
    [
        ("76t58-held-paragraph-centred", "center", None),
        ("76t68-host-right-held-centred", "center", "right"),
        ("76t69-host-right-held-plain", "right", "right"),
    ],
)
def test_the_two_paragraphs_take_the_properties_word_gives_them(stem: str, joined: str, left: str | None) -> None:
    """The mark ending the host carries the held ``w:p``'s, and what's left carries the host's."""

    def jc(paragraph: Any) -> str | None:
        found = paragraph.p_pr.jc if paragraph.p_pr is not None else None
        return found.val.value if found is not None else None

    for pkg in (load(source(stem)), word(stem)):
        first, empty = pkg.main_document_part.contents.body.content[1:3]
        assert (jc(first), jc(empty)) == (joined, left)
        assert not text_of(empty)


def test_a_held_paragraphs_properties_replace_the_hosts() -> None:
    root = tree(f'<w:body><w:p><w:pPr><w:jc w:val="right"/></w:pPr>{A}<w:madeUp><w:p><w:pPr><w:jc w:val="center"/></w:pPr>{B}</w:p></w:madeUp></w:p></w:body>')
    assert read_through.repair(root) == 1
    assert inner(root) == f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>{A}{B}</w:p><w:p><w:pPr><w:jc w:val="right"/></w:pPr></w:p>'


def test_a_host_keeps_its_properties_where_the_held_paragraph_has_none() -> None:
    root = tree(f'<w:body><w:p><w:pPr><w:jc w:val="right"/></w:pPr>{A}<w:madeUp><w:p>{B}</w:p></w:madeUp></w:p></w:body>')
    assert read_through.repair(root) == 1
    assert inner(root) == f'<w:p><w:pPr><w:jc w:val="right"/></w:pPr>{A}{B}</w:p><w:p><w:pPr><w:jc w:val="right"/></w:pPr></w:p>'


def test_what_is_left_of_the_host_doesnt_repeat_its_section_break() -> None:
    root = tree(f'<w:body><w:p><w:pPr><w:jc w:val="right"/><w:sectPr/></w:pPr>{A}<w:madeUp><w:p>{B}</w:p></w:madeUp></w:p></w:body>')
    assert read_through.repair(root) == 1
    assert inner(root) == f'<w:p><w:pPr><w:jc w:val="right"/><w:sectPr/></w:pPr>{A}{B}</w:p><w:p><w:pPr><w:jc w:val="right"/></w:pPr></w:p>'


def test_a_section_break_moves_with_the_properties_a_held_paragraph_replaces() -> None:
    root = tree(f'<w:body><w:p><w:pPr><w:sectPr/></w:pPr>{A}<w:madeUp><w:p><w:pPr><w:jc/></w:pPr>{B}</w:p></w:madeUp></w:p></w:body>')
    assert read_through.repair(root) == 1
    assert inner(root) == f"<w:p><w:pPr><w:jc/></w:pPr>{A}{B}</w:p><w:p><w:pPr><w:sectPr/></w:pPr></w:p>"


def test_a_held_paragraph_at_depth_ends_the_host_too() -> None:
    root = tree(f"<w:body><w:p>{A}<w:madeUp><w:madeUp><w:p>{B}</w:p></w:madeUp></w:madeUp></w:p></w:body>")
    assert read_through.repair(root) == 2
    assert inner(root) == f"<w:p>{A}{B}</w:p><w:p/>"


@pytest.mark.parametrize(
    "xml",
    [
        f"<w:p>{A}<w:madeUp><w:p>{B}</w:p></w:madeUp>{C}</w:p>",  # a run after it: 76t53
        f"<w:p><w:madeUp><w:p>{B}</w:p></w:madeUp>{C}</w:p>",  # first: 76t54
        f"<w:p><w:madeUp><w:p>{B}</w:p></w:madeUp></w:p>",  # alone: 76t55
        f"<w:p>{A}<w:madeUp><w:madeUp><w:p>{B}</w:p></w:madeUp>{C}</w:madeUp></w:p>",  # a run after it, a level down
    ],
)
def test_a_shape_word_refuses_is_left_to_the_second_parse(xml: str) -> None:
    root = tree(f"<w:body>{xml}</w:body>")
    read_through.repair(root)
    assert len(root) == 1 and root[0].find(W + "p") is not None  # the host stands alone, the w:p still in it


def test_a_held_paragraph_in_a_hyperlink_is_beyond_the_wd_and_reads_as_a_gap() -> None:
    """``76t71``: Word keeps the words, and makes the link a complex field across two paragraphs. REQ-316 stops at a ``w:p``."""
    pkg = load(source(IN_A_HYPERLINK))
    assert not any("HIDDEN WORDS" in token for token in tokens(pkg)["body"][2:])
    assert any("HIDDEN WORDS" in token for token in tokens(word(IN_A_HYPERLINK))["body"][2:])
    assert [item.qname for item in pkg.main_document_part.skipped] == [W + "p", MADE_UP]


@pytest.mark.parametrize(("stem", "held"), HELD_REFUSED.items())
def test_a_held_paragraph_in_a_file_word_refuses_is_dropped_and_listed(stem: str, held: int) -> None:
    pkg = load(source(stem))
    assert pkg.main_document_part.contents is not None
    assert sorted(entry(item) for item in pkg.main_document_part.skipped) == sorted([(W + "p", W + "p")] * held + [(MADE_UP, W + "p")])
    assert not any("HIDDEN WORDS" in token for token in tokens(pkg)["body"][2:])  # the first paragraph quotes them


def test_a_paragraph_held_in_another_host_is_only_read_through() -> None:
    got, _pkg = body_tokens(f"<w:madeUp><w:p>{A}</w:p></w:madeUp>")
    assert got == ["t:a", "¶"]


def test_a_w_body_holding_an_element_reads_it_in_order() -> None:
    got, _pkg = body_tokens(f"<w:p>{A}</w:p><w:madeUp><w:p>{B}</w:p><w:p>{C}</w:p></w:madeUp><w:p>{A}</w:p>")
    assert got == ["t:a", "¶", "t:b", "¶", "t:c", "¶", "t:a", "¶"]


def test_each_of_several_is_listed_in_document_order() -> None:
    _got, pkg = body_tokens(f"<w:p><w:madeUp>{A}</w:madeUp></w:p><w:madeUp><w:p><w:other>{B}</w:other></w:p></w:madeUp>")
    assert names(pkg) == [("madeUp", "p"), ("madeUp", "body"), ("other", "p")]


# ---------------------------------------------------------------------------
# a run in a paragraph's w:pPr (Q5, core-ts's D8)
# ---------------------------------------------------------------------------


def test_only_the_run_leaves_a_w_ppr_the_rest_goes_as_a_property_does() -> None:
    got, pkg = body_tokens(f"<w:p><w:pPr><w:madeUp>{A}<w:jc w:val='center'/></w:madeUp></w:pPr>{B}</w:p>")
    assert got == ["t:ab", "¶"]
    (item,) = pkg.main_document_part.skipped
    assert item.qname == MADE_UP  # the fork's own entry, the element still there, listed once
    paragraph = pkg.main_document_part.contents.body.content[0]
    assert paragraph.p_pr.jc is None


def test_text_left_behind_is_absent_and_the_element_goes() -> None:
    """REQ-314: text holds nothing up."""
    root = tree(f"<w:p><w:pPr><w:madeUp>stray{A}more</w:madeUp></w:pPr></w:p>")
    dropped: list[Any] = []
    assert read_through.repair(root, dropped) == 1
    assert inner(root) == f"<w:pPr/>{A}"
    assert len(dropped) == 1


def test_runs_keep_their_order_at_the_head() -> None:
    got, pkg = body_tokens(f"<w:p><w:pPr><w:madeUp>{A}{B}</w:madeUp><w:madeUp>{C}</w:madeUp></w:pPr>{R.format('d')}</w:p>")
    assert got == ["t:abcd", "¶"]
    assert names(pkg) == [("madeUp", "pPr"), ("madeUp", "pPr")]


def test_the_head_is_straight_after_the_w_ppr() -> None:
    root = tree(f"<w:p><w:pPr><w:jc/><w:madeUp>{A}</w:madeUp></w:pPr><w:bookmarkStart/>{B}</w:p>")
    assert read_through.repair(root) == 1
    assert inner(root) == f"<w:pPr><w:jc/></w:pPr>{A}<w:bookmarkStart/>{B}"


def test_a_w_ppr_with_only_a_property_is_untouched() -> None:
    root = tree("<w:p><w:pPr><w:madeUp><w:jc/></w:madeUp></w:pPr></w:p>")
    assert read_through.repair(root) == 0


@pytest.mark.parametrize(
    "xml",
    [
        f"<w:p><w:pPr><w:madeUp><w:madeUp>{A}</w:madeUp></w:madeUp></w:pPr></w:p>",
        f"<w:p><w:pPr><w:rPr><w:madeUp>{A}</w:madeUp></w:rPr></w:pPr></w:p>",
        f"<w:p><w:r><w:rPr><w:madeUp>{A}</w:madeUp></w:rPr></w:r></w:p>",
        f"<w:style><w:pPr><w:madeUp>{A}</w:madeUp></w:pPr></w:style>",
        f"<w:lvl><w:pPr><w:madeUp>{A}</w:madeUp></w:pPr></w:lvl>",
        f"<w:pPrChange><w:pPr><w:madeUp>{A}</w:madeUp></w:pPr></w:pPrChange>",
        f"<w:p><w:pPr><w:pPrChange><w:pPr><w:rPr><w:madeUp>{A}</w:madeUp></w:rPr></w:pPr></w:pPrChange></w:pPr></w:p>",
        f"<w:p><w:pPr><w:pPrChange><w:pPr><w:madeUp><w:madeUp>{A}</w:madeUp></w:madeUp></w:pPr></w:pPrChange></w:pPr></w:p>",
        f"<w:p><w:pPr><w:sectPr><w:madeUp><w:madeUp>{A}</w:madeUp></w:madeUp></w:sectPr></w:pPr></w:p>",
        f"<w:body><w:sectPr><w:madeUp>{A}</w:madeUp></w:sectPr></w:body>",
        f"<w:style><w:pPr><w:sectPr><w:madeUp>{A}</w:madeUp></w:sectPr></w:pPr></w:style>",
        f"<w:style><w:pPr><w:pPrChange><w:pPr><w:madeUp>{A}</w:madeUp></w:pPr></w:pPrChange></w:pPr></w:style>",
    ],
)
def test_only_a_direct_run_in_one_of_a_paragraphs_three_places_moves(xml: str) -> None:
    root = tree(xml)
    assert read_through.repair(root) == 0


def test_a_run_under_a_w_ppr_change_is_read_and_the_change_stays_as_held() -> None:
    """``76t40``, core-ts's D9: Word writes "Unknown" over the author, which is no reading of the text."""
    stem = "76t40-run-in-pPrChange"
    pkg = load(source(stem))
    assert pkg.body.paragraphs[1].text == "HIDDEN WORDSTARGET TEXT"
    assert word(stem).body.paragraphs[1].text == "HIDDEN WORDSTARGET TEXT"
    xml = entries(saved(pkg))["word/document.xml"]
    assert b"<w:pPrChange" in xml and b'w:author="Check 76"' in xml


def test_a_run_in_a_paragraphs_w_sect_pr_is_read_and_the_w_sect_pr_goes() -> None:
    """``76t50``, core-ts's D10: Word's section break goes with it, and so does ours."""
    stem = "76t50-run-in-sectPr"
    pkg = load(source(stem))
    assert pkg.main_document_part.contents is not None
    xml = entries(saved(pkg))["word/document.xml"]
    theirs = entries((CHECK76 / f"{stem}-word365.docx").read_bytes())["word/document.xml"]
    assert xml.count(b"<w:sectPr") == theirs.count(b"<w:sectPr") == 1  # the body's alone
    assert pkg.body.paragraphs[1].text == word(stem).body.paragraphs[1].text


def test_a_w_sect_pr_goes_though_its_element_held_something_else() -> None:
    root = tree(f"<w:body><w:p><w:pPr><w:sectPr><w:madeUp>{A}<w:jc/></w:madeUp></w:sectPr></w:pPr></w:p></w:body>")
    dropped: list[Any] = []
    assert read_through.repair(root, dropped) == 1
    assert inner(root) == f"<w:p><w:pPr/>{A}</w:p>"
    assert [item.qname for item in dropped] == [MADE_UP, W + "sectPr"]


def test_a_w_sect_pr_with_only_a_property_stays() -> None:
    root = tree("<w:p><w:pPr><w:sectPr><w:madeUp><w:jc/></w:madeUp></w:sectPr></w:pPr></w:p>")
    assert read_through.repair(root) == 0


def test_a_w_sect_pr_holding_two_elements_is_listed_once() -> None:
    root = tree(f"<w:body><w:p><w:pPr><w:sectPr><w:madeUp>{A}</w:madeUp><w:madeUp>{B}</w:madeUp></w:sectPr></w:pPr></w:p></w:body>")
    dropped: list[Any] = []
    assert read_through.repair(root, dropped) == 2
    assert [item.qname for item in dropped] == [MADE_UP, W + "sectPr", MADE_UP]
    assert inner(root) == f"<w:p><w:pPr/>{A}{B}</w:p>"


def test_runs_from_all_three_places_keep_document_order() -> None:
    root = tree(
        f"<w:body><w:p><w:pPr><w:madeUp>{A}</w:madeUp><w:sectPr><w:madeUp>{B}</w:madeUp></w:sectPr>"
        f"<w:pPrChange><w:pPr><w:madeUp>{C}</w:madeUp></w:pPr></w:pPrChange></w:pPr></w:p></w:body>"
    )
    assert read_through.repair(root) == 3
    assert inner(root) == f"<w:p><w:pPr><w:pPrChange><w:pPr/></w:pPrChange></w:pPr>{A}{B}{C}</w:p>"


def test_both_rules_in_one_paragraph() -> None:
    got, pkg = body_tokens(f"<w:p><w:pPr><w:madeUp>{A}</w:madeUp></w:pPr>{B}<w:madeUp>{C}</w:madeUp></w:p>")
    assert got == ["t:abc", "¶"]
    assert names(pkg) == [("madeUp", "pPr"), ("madeUp", "p")]


def test_a_paragraph_read_through_with_a_run_in_its_own_w_ppr() -> None:
    got, pkg = body_tokens(f"<w:madeUp><w:p><w:pPr><w:madeUp>{A}</w:madeUp></w:pPr>{B}</w:p></w:madeUp>")
    assert got == ["t:ab", "¶"]
    assert names(pkg) == [("madeUp", "body"), ("madeUp", "pPr")]


# ---------------------------------------------------------------------------
# beside CR-030's pass
# ---------------------------------------------------------------------------


def test_a_w_br_held_in_a_w_p_is_rule_1s_to_place() -> None:
    """Section 3.2's order. A design choice: no Word run holds the shape."""
    got, pkg = body_tokens(f"<w:p>{A}<w:madeUp><w:br/>{B}</w:madeUp></w:p>")
    assert got == ["t:a", "br", "t:b", "¶"]
    part = pkg.main_document_part
    assert names(pkg) == [("madeUp", "p")]
    assert [(w.breaks, w.text_elements) for w in part.warnings] == [(1, 0)]


def test_a_w_br_held_in_w_body_is_rule_3s() -> None:
    got, pkg = body_tokens(f"<w:madeUp><w:br/></w:madeUp><w:p>{A}</w:p>")
    assert got == ["br", "t:a", "¶"]
    assert [(w.breaks, w.text_elements) for w in pkg.main_document_part.warnings] == [(1, 0)]


def test_a_text_element_in_what_it_held_is_rule_2s() -> None:
    got, pkg = body_tokens('<w:p><w:madeUp><w:r><w:t>a<w:madeUp>b</w:madeUp>c</w:t></w:r></w:madeUp></w:p>')
    assert got == ["t:ac", "¶"]
    assert names(pkg) == [("madeUp", "p"), ("madeUp", "t")]


def test_a_part_with_both_shapes_is_repaired_once(monkeypatch: pytest.MonkeyPatch) -> None:
    from docx4j_py import runtime

    parses = []
    real = runtime.parser

    def counting(**kwargs: Any) -> Any:
        parses.append(1)
        return real(**kwargs)

    pkg = package_of(f"<w:p><w:br/><w:madeUp>{A}</w:madeUp></w:p>")
    part = pkg.main_document_part
    monkeypatch.setattr(runtime, "parser", counting)
    assert part.contents is not None
    assert len(parses) == 2
    assert [(w.breaks, w.text_elements) for w in part.warnings] == [(1, 0)]
    assert names(pkg) == [("madeUp", "p")]


def test_the_repair_counts_each_pass() -> None:
    data = f'<w:p {NS}><w:madeUp>{A}</w:madeUp><w:br/></w:p>'.encode()
    repaired, elements, found, texts, _joined = read_through.repair_bytes(data)
    assert (elements, found, texts) == (1, 1, 0)
    assert repaired is not data and repaired.startswith(b"<?xml")


def test_repair_bytes_returns_the_bytes_themselves_where_nothing_changes() -> None:
    data = f"<w:p {NS}>{A}</w:p>".encode()
    assert read_through.repair_bytes(data) == (data, 0, 0, 0, 0)
    assert read_through.repair_bytes(data)[0] is data


def test_a_part_the_pass_leaves_alone_isnt_parsed_again(monkeypatch: pytest.MonkeyPatch) -> None:
    """``76t5``: the report names the element, the pass takes nothing, and the first parse stands."""
    from docx4j_py import runtime

    parses = []
    real = runtime.parser

    def counting(**kwargs: Any) -> Any:
        parses.append(1)
        return real(**kwargs)

    pkg = load(source("76t5-madeup-holding-jc-in-pPr"))
    monkeypatch.setattr(runtime, "parser", counting)
    assert pkg.main_document_part.contents is not None
    assert len(parses) == 1


def test_a_clean_part_runs_no_pass(monkeypatch: pytest.MonkeyPatch) -> None:
    def never(*_args: Any, **_kwargs: Any) -> Any:
        raise AssertionError("a pass ran on a clean part")

    pkg = load(SAMPLES / "Symbols.docx")
    monkeypatch.setattr(read_through, "triggers", never)
    monkeypatch.setattr(read_through, "repair_bytes", never)
    monkeypatch.setattr(breaks, "triggers", never)
    assert pkg.main_document_part.contents is not None
    assert pkg.main_document_part.skipped == []


# ---------------------------------------------------------------------------
# mc:AlternateContent
# ---------------------------------------------------------------------------

ALTERNATE = (
    '<mc:AlternateContent><mc:Choice Requires="wps"><w:p><w:madeUp>{}</w:madeUp></w:p></mc:Choice>'
    "<mc:Fallback><w:p>{}</w:p></mc:Fallback></mc:AlternateContent>"
)


@pytest.mark.parametrize("options", [{}, {"mc_load_mode": "word"}, {"mce_preprocess": True}])
def test_a_branchs_paragraph_reads_as_any_paragraph(options: dict[str, Any]) -> None:
    """A branch's ``w:p`` is typed, so the report names what it holds, in each of the three loads."""
    _got, pkg = body_tokens(f"<w:p>{C}</w:p>" + ALTERNATE.format(A, B), **options)
    assert text(pkg).strip() == "c\na"
    assert names(pkg) == [("madeUp", "p")]
    xml = entries(saved(pkg))["word/document.xml"]
    assert b"<w:madeUp" not in xml
    assert (b"<mc:Fallback>" in xml) is (not options)  # preserve keeps every branch


TAKEN = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" Requires="w14"'
NOT_TAKEN = 'Requires="ext"'


def branch(choice: str, fallback: str, requires: str = TAKEN) -> str:
    return f"<mc:AlternateContent><mc:Choice {requires}>{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"


MADE = f"<w:madeUp>{A}</w:madeUp>"


@pytest.mark.parametrize("options", [{}, {"mc_load_mode": "word"}, {"mce_preprocess": True}])
def test_an_element_in_the_choice_word_takes_is_read_through(options: dict[str, Any]) -> None:
    """REQ-313, ``76t44``. A wildcard holds it, so no report names it."""
    pkg = package_of(f"<w:p>{B}{branch(MADE, C)}</w:p>", **options)
    assert text(pkg).strip() == "ba"


def test_an_element_in_the_fallback_word_takes_is_read_through() -> None:
    """``76t45``: a Choice Word doesn't take."""
    pkg = package_of(f"<w:p>{B}{branch(C, MADE, NOT_TAKEN)}</w:p>")
    assert text(pkg).strip() == "ba"


def test_only_the_branch_taken_is_read_through() -> None:
    root = tree(f"<w:p>{branch(MADE, MADE)}</w:p>")
    assert read_through.repair(root) == 1
    assert f"<mc:Choice {TAKEN}>{A}</mc:Choice><mc:Fallback>{MADE}</mc:Fallback>" in inner(root)


def test_the_choice_not_taken_keeps_its_element() -> None:
    root = tree(f"<w:p>{branch(MADE, MADE, NOT_TAKEN)}</w:p>")
    assert read_through.repair(root) == 1
    assert f"<mc:Choice {NOT_TAKEN}>{MADE}</mc:Choice><mc:Fallback>{A}</mc:Fallback>" in inner(root)


@pytest.mark.parametrize("xml", [f"<w:r>{branch(MADE, C)}</w:r>", f"<w:body>{branch(MADE, C)}</w:body>", f"<w:tc>{branch(MADE, C)}</w:tc>"])
def test_a_branch_outside_a_w_p_isnt_measured(xml: str) -> None:
    assert read_through.repair(tree(xml)) == 0


@pytest.mark.parametrize(
    ("xml", "may_hold"),
    [
        (f'<w:p xmlns:w="{W_URI}"><w:r/></w:p>', False),
        (f'<w:p xmlns:w="{W_URI}"><w:madeUp/></w:p>', False),  # no AlternateContent: the report's business
        (f'<w:p xmlns:w="{W_URI}" xmlns:mc="urn:mc"><mc:AlternateContent><w:r/></mc:AlternateContent></w:p>', False),
        (f'<w:p xmlns:w="{W_URI}" xmlns:mc="urn:mc"><mc:AlternateContent><w:madeUp/></mc:AlternateContent></w:p>', True),
        (f'<x:p xmlns:x="{W_URI}" xmlns:mc="urn:mc"><mc:AlternateContent><x:madeUp/></mc:AlternateContent></x:p>', True),
        (f'<x:p xmlns:x="{W_URI}" xmlns:mc="urn:mc"><mc:AlternateContent><x:r/></mc:AlternateContent></x:p>', False),
        ('<w:p xmlns:w="urn:not-it" xmlns:mc="urn:mc"><mc:AlternateContent><w:madeUp/></mc:AlternateContent></w:p>', False),
    ],
)
def test_what_marks_a_part_for_the_branch_pass(xml: str, may_hold: bool) -> None:
    assert read_through.in_a_branch(xml.encode()) is may_hold


def test_a_part_with_branches_and_no_unknown_name_runs_no_pass(monkeypatch: pytest.MonkeyPatch) -> None:
    def never(*_args: Any, **_kwargs: Any) -> Any:
        raise AssertionError("a pass ran on a clean part")

    pkg = package_of(f"<w:p>{branch(A, B)}</w:p>")
    monkeypatch.setattr(read_through, "repair_bytes", never)
    assert pkg.main_document_part.contents is not None


def test_text_an_element_holds_directly_is_absent() -> None:
    """REQ-314, ``76t41``."""
    got, pkg = body_tokens(f"<w:p>{A}<w:madeUp>stray</w:madeUp>{B}</w:p>")
    assert got == ["t:ab", "¶"]
    assert names(pkg) == [("madeUp", "p")]


# ---------------------------------------------------------------------------
# strict, fragments and the source tree
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("stem", AS_WORD)
def test_strict_refuses_though_no_words_go(stem: str) -> None:
    """Q1: the tags are lost whatever the words do, so ``strict=True`` says so."""
    pkg = load(source(stem), options=LoadOptions(strict=True))
    with pytest.raises(Docx4JException, match="strict=True"):
        assert part_of(pkg, AS_WORD[stem][0]).contents is not None


@pytest.mark.parametrize(
    "fragment",
    [f"<w:p><w:madeUp>{A}</w:madeUp></w:p>", f"<w:madeUp><w:p>{A}</w:p></w:madeUp>", f"<w:p><w:pPr><w:madeUp>{A}</w:madeUp></w:pPr></w:p>"],
)
def test_insert_ooxml_refuses_such_a_fragment(fragment: str) -> None:
    """Q2: refusing what a read can't keep is that route's rule."""
    pkg = create_package()
    before = len(pkg.body.paragraphs)
    with pytest.raises(FragmentError, match="1 item"):
        pkg.body.insert_ooxml(fragment)
    assert len(pkg.body.paragraphs) == before


def test_a_lenient_fragment_is_read_through() -> None:
    (paragraph,) = wml.all(f"<w:p><w:madeUp>{A}</w:madeUp></w:p>", wrapper="body", lenient=True, _breaks=[])
    assert text_of(paragraph) == "a"


def test_a_lenient_fragment_in_a_body_reads_a_top_level_element() -> None:
    first, second = wml.all(f"<w:madeUp><w:p>{A}</w:p><w:p>{B}</w:p></w:madeUp>", wrapper="body", lenient=True, _breaks=[])
    assert (text_of(first), text_of(second)) == ("a", "b")


def test_a_lenient_fragment_keeps_its_break_record() -> None:
    found: list[Any] = []
    (paragraph,) = wml.all(f"<w:p><w:madeUp><w:br/>{A}</w:madeUp></w:p>", wrapper="body", lenient=True, _breaks=found)
    assert [(record.breaks, record.text_elements) for record in found] == [(1, 0)]
    assert text_of(paragraph) == "\na"


def test_wml_stays_strict() -> None:
    with pytest.raises(FragmentError):
        wml(f"<w:p><w:madeUp>{A}</w:madeUp></w:p>")


def test_a_fragment_holding_a_paragraph_in_a_paragraph_is_refused_then_joined() -> None:
    fragment = f"<w:p>{A}<w:madeUp><w:p>{B}</w:p></w:madeUp></w:p>"
    with pytest.raises(FragmentError, match="2 item"):
        create_package().body.insert_ooxml(fragment)
    first, second = wml.all(fragment, wrapper="body", lenient=True, _breaks=[])
    assert (text_of(first), text_of(second)) == ("ab", "")


def test_a_lenient_wml_runs_no_pass() -> None:
    """Code's own fragments aren't files: ``wml`` drops as it always has."""
    paragraph = wml(f"<w:p><w:madeUp>{A}</w:madeUp>{B}</w:p>", lenient=True)
    assert text_of(paragraph) == "b"


def test_the_source_tree_is_read_through_too() -> None:
    """The save's check compares with this, so what was read isn't taken for the session's."""
    part = load(source("76t-madeup-holding-a-run")).main_document_part
    assert "HIDDEN WORDS" in text_of(part.source_tree())
    assert part.skipped == []  # the part itself isn't touched


def test_set_bytes_clears_the_report() -> None:
    pkg = package_of(f"<w:p><w:madeUp>{A}</w:madeUp></w:p>")
    part = pkg.main_document_part
    assert part.contents is not None and part.skipped
    part.set_bytes(entries(saved(create_package()))["word/document.xml"])
    assert part.contents is not None
    assert part.skipped == []
