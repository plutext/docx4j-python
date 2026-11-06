"""CR-039 Phase B: a main part's root, and a ``w:body`` given twice, read as Word reads them.

The oracle is the Word bridge's files (``tests/fixtures/repeated/``): each source with Word's
save and its text save. WD REQ-317 to REQ-321 are the rules, proposed at 0.96.
"""

from __future__ import annotations

import hashlib
import io
import logging
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))
import canon  # noqa: E402

from docx4j_py import load  # noqa: E402
from docx4j_py.model.content.body import body_of  # noqa: E402
from docx4j_py.model.content.errors import ContentError  # noqa: E402
from docx4j_py.model.text import TextWarning, extract, plain_text  # noqa: E402
from docx4j_py.openpackaging import LoadOptions, read_through, repeated  # noqa: E402
from docx4j_py.openpackaging.exceptions import Docx4JException  # noqa: E402
from docx4j_py.openpackaging.parts import xml_part  # noqa: E402
from docx4j_py.traversal import misplaced  # noqa: E402

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "repeated"
SAMPLES = Path(__file__).resolve().parents[2] / "samples"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W = f"{{{W_NS}}}"
DOC = "/word/document.xml"

#: Word read every word of these, with a section of its own.
READ = [
    "rc-10-body",
    "rc-11-sdtContent",
    "mr-01-madeup-p",
    "mr-10-two-p",
    "mr-11-p-sectpr",
    "mr-12-tbl",
    "mr-13-sdt",
    "mr-15-body",
    "mr-17-nested",
    "mr-18-p-properties",
    "mr-30-sdtcontent-p",
    "mr-32-p-root",
    "mr-40-branch-madeup-p",
    "mr-50-madeup-p-styles",
    "nd-20-body-only-child",
    "nd-21-body-middle-no-sectPr",
    "nd-22-body-middle-with-sectPr",
    "nd-23-body-last-with-sectPr",
    "nd-24-body-three-deep",
    "rt-15-body-section-break",
    "rt-16-body-second-no-sectPr",
    "rt-17-body-third-no-sectPr",
    "rw-20-mc-root-choice",
    "rw-21-mc-root-fallback",
    "rw-22-mc-root-fallback-only",
    "rw-23-mc-root-two-documents",
    "rw-26-mc-root-two-in-fallback",
    "rw-27-mc-root-two-in-choice-sizes",
    "or-31-madeup",
    "or-32-body",
]
REFUSED = [
    "rw-10-foo-two-documents",
    "rw-11-zfoo-two-documents",
    "rw-12-wfoo-two-documents",
    "rw-13-foo-one-document",
    "rw-30-word2003-root",
    "rw-31-flat-opc-root",
    "or-12-footnotes",
    "or-13-endnotes",
    "or-15-footnotes-bare",
    "mr-03-foo-document",
    "mr-14-run",
    "mr-20-other-namespace-p",
    "mr-21-no-namespace-p",
    "mr-31-tc-p",
]
#: Each with the root left as it is, and what the root held.
UNEXPECTED = {
    "or-10-hdr": ("w:hdr", ["p"]),
    "or-11-ftr": ("w:ftr", ["p"]),
    "or-14-comments": ("w:comments", ["comment"]),
    "or-30-styles": ("w:styles", ["style"]),
    "mr-22-ignorable-p": ("{urn:docx4j:check79}madeUp", ["p"]),
}
#: A property given twice: where it stands, and what Word's save holds of the target paragraph's.
PROPERTIES = ["rc-20-pPr-values", "rc-21-pPr-merge", "rc-22-rPr-values", "rc-23-rPr-merge", "rc-30-jc", "rc-33-sz"]
NO_MAIN_PART = ["or-40-hdr-type-header", "or-41-document-type-header"]


def source(name: str) -> bytes:
    return (FIXTURES / f"{name}.docx").read_bytes()


def part_of(data: bytes) -> bytes:
    return zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml")


def word_save(name: str) -> bytes:
    return part_of((FIXTURES / f"{name}-word365.docx").read_bytes())


def word_text(name: str) -> str:
    return (FIXTURES / f"{name}-word365.txt").read_text(encoding="utf-8").strip()


def saved(package: Any, **options: Any) -> bytes:
    out = io.BytesIO()
    package.save(out, **options)
    return out.getvalue()


def with_part(xml: bytes, name: str = "or-00-control") -> bytes:
    """The control's package holding `xml` as its main part."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source(name))) as zin, zipfile.ZipFile(out, "w") as zout:
        for info in zin.infolist():
            zout.writestr(info, xml if info.filename == "word/document.xml" else zin.read(info.filename))
    return out.getvalue()


def local_names(items: list[Any]) -> list[str]:
    return [etree.QName(item.qname).localname for item in items]


def page_widths(xml: bytes) -> list[str]:
    return [size.get(f"{W}w") for size in etree.fromstring(xml).iter(f"{W}pgSz")]


# --- the fixtures ----------------------------------------------------------------


def test_every_fixture_is_the_file_word_read() -> None:
    listed = dict(reversed(line.split()) for line in (FIXTURES / "SHA1SUMS").read_text().splitlines())
    on_disk = {path.name for path in FIXTURES.iterdir()} - {"SHA1SUMS", "README.md"}
    assert set(listed) == on_disk
    for name, digest in listed.items():
        assert hashlib.sha1((FIXTURES / name).read_bytes()).hexdigest() == digest, name


@pytest.mark.parametrize("name", ["rc-00-control", "rw-00-control", "rw-01-control-mc-choice", "or-00-control"])
def test_a_control_is_read_as_it_was_and_warned_of_by_nothing(name: str) -> None:
    package = load(source(name))
    part = package.main_document_part
    assert plain_text(extract(package)).strip() == word_text(name)
    assert not [w for w in part.warnings if isinstance(w, repeated.ReadAsWord | repeated.UnexpectedRoot)]
    assert not part.skipped


# --- what Word reads -------------------------------------------------------------


@pytest.mark.parametrize("name", READ)
def test_the_words_are_words_word_read(name: str) -> None:
    assert plain_text(extract(load(source(name)))).strip() == word_text(name)


@pytest.mark.parametrize("name", READ)
def test_a_save_has_word_s_words_and_section(name: str) -> None:
    package = load(source(name))
    package.main_document_part.contents
    ours = part_of(saved(package))
    theirs = word_save(name)
    assert etree.fromstring(ours).tag == f"{W}document"

    def words(xml: bytes) -> list[str]:
        body = etree.fromstring(xml).find(f"{W}body")
        return [text for paragraph in body.iter(f"{W}p") if (text := "".join(paragraph.itertext()))]

    assert words(ours) == words(theirs)
    if page_widths(ours):  # a file that states no section has none, where Word's save writes a default
        assert page_widths(ours) == page_widths(theirs)


@pytest.mark.parametrize("name", READ)
def test_a_save_is_read_again_with_nothing_to_say(name: str) -> None:
    package = load(source(name))
    package.main_document_part.contents
    again = load(saved(package))
    assert plain_text(extract(again)).strip() == word_text(name)
    assert not again.main_document_part.skipped
    assert not again.warnings


@pytest.mark.parametrize("name", READ)
def test_body_writes_land_in_the_part(name: str) -> None:
    package = load(source(name))
    body_of(package).insert_paragraph("NEW WORDS")
    assert b"NEW WORDS" in part_of(saved(package))


@pytest.mark.parametrize(
    "options", [LoadOptions(mce_preprocess=True), LoadOptions(mc_load_mode="word")], ids=["mce_preprocess", "word"]
)
@pytest.mark.parametrize("name", ["rw-20-mc-root-choice", "rw-21-mc-root-fallback", "rw-23-mc-root-two-documents"])
def test_the_mc_load_options_read_the_same(name: str, options: LoadOptions) -> None:
    assert plain_text(extract(load(source(name), options=options))).strip() == word_text(name)


def test_nothing_selected_is_an_empty_document_with_a_body() -> None:
    package = load(source("rw-24-mc-root-nothing-selected"))
    part = package.main_document_part
    assert plain_text(extract(package)).strip() == word_text("rw-24-mc-root-nothing-selected") == ""
    assert local_names(part.skipped) == ["Choice"]
    assert [type(warning) for warning in part.warnings] == [repeated.ReadAsWord]
    body_of(package).insert_paragraph("NEW WORDS")
    assert b"NEW WORDS" in part_of(saved(package))


def test_a_prefix_only_the_mc_root_declares_survives_the_save() -> None:
    """The part's source declarations are the read root's, so ``mc:Ignorable``'s prefix is still declared."""
    xml = (
        f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}" xmlns:x="urn:x"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><mc:Choice Requires="w14">'
        '<w:document mc:Ignorable="x"><w:body><w:p><w:r><w:t>KEPT</w:t></w:r></w:p></w:body></w:document>'
        "</mc:Choice></mc:AlternateContent>"
    ).encode()
    package = load(with_part(xml))
    package.main_document_part.contents
    root = etree.fromstring(part_of(saved(package)))
    assert root.get(f"{{{MC_NS}}}Ignorable") == "x"
    assert root.nsmap.get("x") == "urn:x"


def test_a_prefix_only_the_branch_declares_survives_the_save() -> None:
    """The part's source declarations are read again from the root the branch gave."""
    xml = (
        f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback>'
        '<w:document xmlns:x="urn:x" mc:Ignorable="x"><w:body><w:p><w:r><w:t>KEPT</w:t></w:r></w:p></w:body></w:document>'
        "</mc:Fallback></mc:AlternateContent>"
    ).encode()
    package = load(with_part(xml))
    package.main_document_part.contents
    root = etree.fromstring(part_of(saved(package)))
    assert root.get(f"{{{MC_NS}}}Ignorable") == "x"
    assert root.nsmap.get("x") == "urn:x"


def test_text_after_the_branch_s_document_stays_out_of_the_part() -> None:
    held = "<w:document><w:body><w:p><w:r><w:t>KEPT</w:t></w:r></w:p></w:body></w:document>stray"
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback>{held}</mc:Fallback></mc:AlternateContent>'
    assert plain_text(extract(load(with_part(xml.encode())))).strip() == "KEPT"


def test_what_a_branch_holds_after_the_root_it_reads_is_listed() -> None:
    held = f'<w:hdr xmlns:w="{W_NS}"><w:p/></w:hdr><w:ftr xmlns:w="{W_NS}"><w:p/></w:ftr>'
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}"><mc:Fallback>{held}</mc:Fallback></mc:AlternateContent>'
    package = load(with_part(xml.encode()))
    package.main_document_part.contents
    assert sorted(local_names(package.main_document_part.skipped)) == ["ftr", "p"]


def test_a_nested_mc_root_is_read_through_each_level() -> None:
    inner = (
        '<mc:AlternateContent><mc:Fallback><w:document><w:body><w:p><w:r><w:t>DEEP</w:t></w:r></w:p></w:body>'
        "</w:document></mc:Fallback></mc:AlternateContent>"
    )
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback>{inner}</mc:Fallback></mc:AlternateContent>'
    assert plain_text(extract(load(with_part(xml.encode())))).strip() == "DEEP"


# --- two bodies ------------------------------------------------------------------


def test_two_bodies_keep_one_section_and_list_the_earlier() -> None:
    package = load(source("rc-10-body"))
    part = package.main_document_part
    body = part.contents.body
    assert len([item for item in body.content]) == 2
    assert body.sect_pr is not None
    assert local_names(part.skipped) == ["sectPr"]
    assert [(w.root, w.children) for w in part.warnings] == [(None, 2)]


def test_a_later_body_with_no_section_leaves_the_first_s_at_the_end() -> None:
    xml = (
        f'<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906"/></w:sectPr>'
        "</w:body><w:body><w:p><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>"
    ).encode()
    package = load(with_part(xml))
    part = package.main_document_part
    root = etree.fromstring(part_of(saved(_read(package))))
    assert [etree.QName(child).localname for child in root.find(f"{W}body")] == ["p", "p", "sectPr"]
    assert not part.skipped


def _read(package: Any) -> Any:
    package.main_document_part.contents
    return package


def test_the_last_section_stated_stands_where_the_last_body_has_none() -> None:
    """``rt-17``: the third body has no ``w:sectPr``, and the second's stands, with the first's listed."""
    name = "rt-17-body-third-no-sectPr"
    package = load(source(name))
    assert page_widths(part_of(saved(_read(package)))) == page_widths(word_save(name)) == ["15840"]
    assert local_names(package.main_document_part.skipped) == ["sectPr"]


def test_a_section_in_a_paragraph_stays_where_it_is() -> None:
    """``rt-15``: the break's ``w:sectPr`` is the paragraph's, and only the bodies' final ones compete."""
    name = "rt-15-body-section-break"
    package = load(source(name))
    root = etree.fromstring(part_of(saved(_read(package))))
    assert [etree.QName(section.getparent()).localname for section in root.iter(f"{W}sectPr")] == ["pPr", "body"]
    assert page_widths(etree.tostring(root)) == page_widths(word_save(name))


def test_three_bodies_join_in_order() -> None:
    bodies = "".join(f"<w:body><w:p><w:r><w:t>{word}</w:t></w:r></w:p></w:body>" for word in "ABC")
    package = load(with_part(f'<w:document xmlns:w="{W_NS}">{bodies}</w:document>'.encode()))
    assert plain_text(extract(package)).split() == ["A", "B", "C"]
    assert [w.children for w in package.main_document_part.warnings] == [2]


def test_what_a_later_document_holds_beside_its_body_is_listed() -> None:
    second = "<w:document><w:background/><w:body><w:p><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>"
    first = "<w:document><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p></w:body></w:document>"
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback>{first}{second}</mc:Fallback></mc:AlternateContent>'
    package = load(with_part(xml.encode()))
    assert plain_text(extract(package)).split() == ["A", "B"]
    assert local_names(package.main_document_part.skipped) == ["background"]


# --- a child given more than once ------------------------------------------------


def target_properties(xml: bytes) -> dict[str, dict[str, str]]:
    """What the target paragraph's ``w:pPr`` and ``w:rPr`` hold: each child's attributes by local name."""
    for paragraph in etree.fromstring(xml).iter(f"{W}p"):
        if "TARGET TEXT" in "".join(paragraph.itertext()):
            return {
                etree.QName(child).localname: {etree.QName(key).localname: value for key, value in child.attrib.items()}
                for properties in paragraph.iter(f"{W}pPr", f"{W}rPr")
                for child in properties
            }
    raise AssertionError("no target paragraph")


@pytest.mark.parametrize("name", PROPERTIES)
def test_a_property_given_twice_is_what_word_s_save_holds(name: str) -> None:
    package = load(source(name))
    assert target_properties(part_of(saved(_read(package)))) == target_properties(word_save(name))


@pytest.mark.parametrize(("name", "value"), [("rc-31-b-on-off", "0"), ("rc-32-b-off-on", None)])
def test_the_last_w_b_stands(name: str, value: str | None) -> None:
    """Word's save of off writes no ``w:b``, and of on a bare one: here the last as written, which reads the same."""
    held = target_properties(part_of(saved(_read(load(source(name))))))
    assert held["b"].get("val") == value


def test_the_later_section_stands_in_one_body() -> None:
    name = "rc-40-sectPr"
    ours = part_of(saved(_read(load(source(name)))))
    assert page_widths(ours) == page_widths(word_save(name)) == ["15840"]


def test_the_later_table_properties_stand() -> None:
    name = "rc-41-tblPr"
    ours = etree.fromstring(part_of(saved(_read(load(source(name))))))
    theirs = etree.fromstring(word_save(name))
    [properties] = ours.iter(f"{W}tblPr")
    assert properties.find(f"{W}jc").get(f"{W}val") == next(theirs.iter(f"{W}tblPr")).find(f"{W}jc").get(f"{W}val") == "right"


def test_what_a_merge_replaced_is_listed_and_what_it_added_isn_t() -> None:
    assert local_names(_read(load(source("rc-20-pPr-values"))).main_document_part.skipped) == ["jc"]
    assert not _read(load(source("rc-21-pPr-merge"))).main_document_part.skipped


def test_a_second_box_stays_as_it_was() -> None:
    """Q6: ``w:txbxContent`` waits for a second shape, so the first stands and the second is listed."""
    package = _read(load(source("rc-12-txbxContent")))
    assert local_names(package.main_document_part.skipped) == ["txbxContent"]


def repaired(xml: str) -> tuple[etree._Element, int]:
    root = etree.fromstring(f'<w:document xmlns:w="{W_NS}"><w:body>{xml}</w:body></w:document>')
    return root, repeated.repair(root)


def test_a_run_level_control_is_left_as_it_is() -> None:
    root, read_as_one = repaired("<w:p><w:sdt><w:sdtContent/><w:sdtContent/></w:sdt></w:p>")
    assert read_as_one == 0
    assert len(list(root.iter(f"{W}sdtContent"))) == 2


def test_a_control_in_a_cell_is_a_block_one() -> None:
    cell = "<w:tc><w:sdt><w:sdtContent><w:p/></w:sdtContent><w:sdtContent><w:p/></w:sdtContent></w:sdt></w:tc>"
    root, read_as_one = repaired(f"<w:tbl><w:tr>{cell}</w:tr></w:tbl>")
    assert read_as_one == 1
    [content] = root.iter(f"{W}sdtContent")
    assert len(content) == 2


def test_a_later_value_of_any_property_wins_a_merge() -> None:
    """Not only the names kept last: a merge replaces whatever the later ``w:pPr`` states."""
    root, read_as_one = repaired('<w:p><w:pPr><w:ind w:left="1"/></w:pPr><w:pPr><w:ind w:left="2"/></w:pPr></w:p>')
    assert read_as_one == 1
    assert [ind.get(f"{W}left") for ind in root.iter(f"{W}ind")] == ["2"]


def test_a_body_in_a_later_body_gives_way_too() -> None:
    bodies = "<w:body><w:p/></w:body><w:body><w:body><w:p/></w:body></w:body>"
    root = etree.fromstring(f'<w:document xmlns:w="{W_NS}">{bodies}</w:document>')
    assert repeated.repair(root) == 2
    assert len(list(root.iter(f"{W}body"))) == 1
    assert len(list(root.iter(f"{W}p"))) == 2


def test_three_copies_read_as_two_do() -> None:
    root, read_as_one = repaired('<w:p><w:pPr><w:jc w:val="left"/><w:jc w:val="center"/><w:jc w:val="right"/></w:pPr></w:p>')
    assert read_as_one == 2
    assert [jc.get(f"{W}val") for jc in root.iter(f"{W}jc")] == ["right"]


def test_a_property_in_a_part_not_the_main_one_is_left_as_it_is() -> None:
    """Every measurement is the main part."""
    root = etree.fromstring(f'<w:hdr xmlns:w="{W_NS}"><w:p><w:pPr><w:jc/><w:jc/></w:pPr></w:p></w:hdr>')
    assert repeated.repair(root) == 0
    assert len(list(root.iter(f"{W}jc"))) == 2


def test_a_part_that_holds_each_child_once_changes_nothing() -> None:
    data = f'<w:document xmlns:w="{W_NS}"><w:body><w:p><w:pPr><w:jc w:val="left"/></w:pPr></w:p></w:body></w:document>'.encode()
    assert read_through.repair_bytes(data)[0] is data


@pytest.mark.parametrize("name", ["jc", "b", "sz", "pPr", "rPr", "tblPr", "sectPr", "sdtContent", "body"])
def test_the_report_triggers_on_each_child_measured(name: str) -> None:
    class Item:
        qname = f"{W}{name}"

    assert repeated.triggers([Item()])


def test_the_report_doesn_t_trigger_on_anything_else() -> None:
    class Item:
        qname = f"{W}txbxContent"

    assert not repeated.triggers([Item()])


def test_a_repeated_property_under_a_root_read_gives_one_warning() -> None:
    held = '<w:document><w:body><w:p><w:pPr><w:jc w:val="left"/><w:jc w:val="right"/></w:pPr></w:p></w:body></w:document>'
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback>{held}</mc:Fallback></mc:AlternateContent>'
    package = _read(load(with_part(xml.encode())))
    assert [(w.root, w.children) for w in package.main_document_part.warnings] == [("mc:AlternateContent", 1)]


# --- an unknown root -------------------------------------------------------------


def test_a_section_directly_in_an_unknown_root_has_no_effect() -> None:
    package = _read(load(source("mr-11-p-sectpr")))
    assert page_widths(part_of(saved(package))) == []
    assert "sectPr" in local_names(package.main_document_part.skipped)


def test_an_empty_unknown_root_is_an_empty_document() -> None:
    package = load(source("mr-16-empty"))
    assert plain_text(extract(package)).strip() == word_text("mr-16-empty") == ""
    body_of(package).insert_paragraph("NEW WORDS")
    assert b"NEW WORDS" in part_of(saved(package))


def test_a_body_in_an_unknown_root_is_the_body_and_what_stands_beside_it_is_listed() -> None:
    xml = f'<w:madeUp xmlns:w="{W_NS}"><w:p/><w:body><w:p><w:r><w:t>WORDS</w:t></w:r></w:p></w:body></w:madeUp>'.encode()
    package = load(with_part(xml))
    assert plain_text(extract(package)).strip() == "WORDS"
    assert local_names(package.main_document_part.skipped) == ["p"]


def test_two_bodies_in_an_unknown_root_join() -> None:
    bodies = "".join(f"<w:body><w:p><w:r><w:t>{word}</w:t></w:r></w:p></w:body>" for word in "AB")
    package = load(with_part(f'<w:madeUp xmlns:w="{W_NS}">{bodies}</w:madeUp>'.encode()))
    assert plain_text(extract(package)).split() == ["A", "B"]


def test_an_unknown_root_holding_nothing_it_reads_is_another_root() -> None:
    package = load(with_part(f'<w:madeUp xmlns:w="{W_NS}"><w:jc/></w:madeUp>'.encode()))
    assert [(w.kind, w.detail) for w in extract(package).warnings] == [("part.unexpected_root", "w:madeUp")]


def test_a_root_understood_holding_a_paragraph_is_another_root() -> None:
    """Unmeasured, so REQ-318's last row doesn't take it: REQ-320's."""
    ns = "http://schemas.microsoft.com/office/word/2010/wordml"
    package = load(with_part(f'<x:foo xmlns:x="{ns}" xmlns:w="{W_NS}"><w:p/></x:foo>'.encode()))
    assert [w.kind for w in extract(package).warnings] == ["part.unexpected_root"]


# --- the roots Word refuses ------------------------------------------------------


@pytest.mark.parametrize("name", REFUSED)
def test_a_root_word_refuses_is_refused_when_read(name: str) -> None:
    package = load(source(name))  # the load stays lazy
    part = package.main_document_part
    with pytest.raises(Docx4JException, match="which Word refuses to open") as refused:
        part.contents
    assert str(part.part_name) in str(refused.value)
    assert part.unreadable is not None
    with pytest.raises(Docx4JException, match="which Word refuses to open"):
        extract(load(source(name)))


@pytest.mark.parametrize("name", REFUSED)
def test_a_root_word_refuses_saves_as_it_came(name: str) -> None:
    data = source(name)
    package = load(data)
    with pytest.raises(Docx4JException):
        package.main_document_part.contents
    assert part_of(saved(package, check=False)) == part_of(data)


@pytest.mark.parametrize("name", REFUSED)
def test_the_validator_reports_a_refused_root_and_goes_on(name: str) -> None:
    findings = misplaced(load(source(name)))
    assert [finding.code for finding in findings if finding.part_name == DOC] == ["part.root_refused"]


def test_a_root_refused_holds_to_what_was_measured() -> None:
    """REQ-318's last clause is an unknown ``w:`` root or a foreign one: a known one is REQ-320's."""
    held = f'<w:document xmlns:w="{W_NS}"><w:body/></w:document>'
    package = load(with_part(f'<w:hdr xmlns:w="{W_NS}">{held}</w:hdr>'.encode()))
    assert plain_text(extract(package)).strip() == ""
    assert [type(w) for w in package.main_document_part.warnings] == [repeated.UnexpectedRoot]


@pytest.mark.parametrize("root", ["w:footnotes", "w:foo"])
def test_a_refusal_under_a_branch_is_a_refusal(root: str) -> None:
    held = f"<{root}><w:document><w:body/></w:document></{root}>"
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback>{held}</mc:Fallback></mc:AlternateContent>'
    with pytest.raises(Docx4JException, match="which Word refuses to open"):
        load(with_part(xml.encode())).main_document_part.contents


def test_xml_not_well_formed_under_another_root_is_refused_with_lxml_s_reason() -> None:
    package = load(with_part(f'<w:hdr xmlns:w="{W_NS}"><w:p></w:hdr>'.encode()))
    with pytest.raises(Docx4JException, match="Problem with part /word/document.xml: Opening and ending tag mismatch"):
        package.main_document_part.contents


# --- a root left as it is --------------------------------------------------------


@pytest.mark.parametrize("name", UNEXPECTED)
def test_another_root_yields_no_text_and_says_so(name: str) -> None:
    root, held = UNEXPECTED[name]
    package = load(source(name))
    result = extract(package)
    assert plain_text(result).strip() == ""
    [warning] = [w for w in result.warnings if w.kind == TextWarning.PART_UNEXPECTED_ROOT]
    assert warning.kind == "part.unexpected_root"
    assert warning.part_name == DOC
    assert warning.detail == root
    assert warning.to_dict() == {"kind": "part.unexpected_root", "address": None, "selector": DOC, "detail": root}
    part = package.main_document_part
    assert local_names(part.skipped) == held
    assert [(w.code, w.root) for w in part.warnings] == [("part.unexpected_root", root)]
    assert [type(w) for w in package.warnings] == [repeated.UnexpectedRoot]


@pytest.mark.parametrize("name", UNEXPECTED)
def test_body_refuses_a_write_into_another_root(name: str) -> None:
    with pytest.raises(ContentError) as refused:
        body_of(load(source(name)))
    assert refused.value.code == "body.no_content"


def test_a_taken_branch_holding_another_root_reads_as_that_root() -> None:
    name = "or-50-mc-hdr-choice"
    package = load(source(name))
    result = extract(package)
    assert plain_text(result).strip() == ""
    assert [w.detail for w in result.warnings if w.kind == TextWarning.PART_UNEXPECTED_ROOT] == ["w:hdr"]
    assert sorted(local_names(package.main_document_part.skipped)) == ["Fallback", "p"]


def test_a_foreign_root_holding_a_body_yields_nothing() -> None:
    """REQ-320: the class alone would have bound its ``w:body``."""
    xml = f'<foo xmlns:w="{W_NS}"><w:body><w:p><w:r><w:t>WORDS</w:t></w:r></w:p></w:body></foo>'.encode()
    package = load(with_part(xml))
    assert plain_text(extract(package)).strip() == ""
    assert local_names(package.main_document_part.skipped) == ["body"]
    assert [w.root for w in package.main_document_part.warnings] == ["foo"]


def test_an_empty_root_still_says_so() -> None:
    package = load(with_part(b"<foo/>"))
    assert [w.kind for w in extract(package).warnings] == ["part.unexpected_root"]


def test_a_namespace_outside_the_model_is_named_in_full() -> None:
    package = load(with_part(b'<x:foo xmlns:x="urn:x"/>'))
    assert [w.detail for w in extract(package).warnings] == ["{urn:x}foo"]


# --- a package with no main part -------------------------------------------------


@pytest.mark.parametrize("name", NO_MAIN_PART)
def test_extract_refuses_a_package_with_no_main_part(name: str) -> None:
    with pytest.raises(Docx4JException, match="no main document part"):
        extract(load(source(name)))


@pytest.mark.parametrize("name", NO_MAIN_PART)
def test_body_of_refuses_a_package_with_no_main_part(name: str) -> None:
    with pytest.raises(ContentError) as refused:
        body_of(load(source(name)))
    assert refused.value.code == "body.no_part"


# --- the report ------------------------------------------------------------------


@pytest.mark.parametrize(
    ("name", "root", "children"),
    [
        ("rw-20-mc-root-choice", "mc:AlternateContent", 0),
        ("rw-23-mc-root-two-documents", "mc:AlternateContent", 2),
        ("or-32-body", "w:body", 0),
        ("or-31-madeup", "w:madeUp", 0),
    ],
)
def test_one_warning_a_part_says_what_was_read(name: str, root: str, children: int) -> None:
    package = load(source(name))
    part = package.main_document_part
    part.contents
    [warning] = part.warnings
    assert isinstance(warning, repeated.ReadAsWord)
    assert (warning.part_name, warning.root, warning.children) == (DOC, root, children)
    assert warning.to_dict()["code"] == "part.read_as_word"
    assert root in warning.message
    assert package.warnings == [warning]


def test_a_mc_root_holding_a_document_of_two_bodies_gives_one_warning() -> None:
    bodies = "<w:body><w:p/></w:body><w:body><w:p/></w:body>"
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback><w:document>{bodies}</w:document></mc:Fallback></mc:AlternateContent>'
    package = load(with_part(xml.encode()))
    package.main_document_part.contents
    assert [(w.root, w.children) for w in package.main_document_part.warnings] == [("mc:AlternateContent", 1)]


def test_the_warning_is_logged_once_a_part(caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level(logging.WARNING, logger="docx4j_py.openpackaging"):
        load(source("or-32-body")).main_document_part.contents
    assert [record.getMessage() for record in caplog.records] == [f"{DOC}: root w:body; read as Word reads them"]


@pytest.mark.parametrize(
    ("name", "loads"),
    [
        ("or-32-body", True),
        ("rc-11-sdtContent", True),
        ("rc-21-pPr-merge", True),
        ("rc-23-rPr-merge", True),
        ("rc-10-body", False),
        ("rc-20-pPr-values", False),
        ("rc-30-jc", False),
        ("rw-20-mc-root-choice", False),
        ("or-31-madeup", False),
    ],
)
def test_strict_loads_a_part_that_lost_nothing(name: str, loads: bool) -> None:
    """A bare ``w:body``, a join and a merge lost nothing. An earlier section, a value, a branch and a tag did."""
    part = load(source(name), options=LoadOptions(strict=True)).main_document_part
    if loads:
        assert part.contents is not None
    else:
        with pytest.raises(Docx4JException, match="were skipped and strict=True"):
            part.contents


# --- the save and the round trip -------------------------------------------------


@pytest.mark.parametrize("name", READ + PROPERTIES)
def test_canon_counts_the_reading_and_nothing_else(name: str) -> None:
    data = source(name)
    package = load(data)
    package.main_document_part.contents
    report = canon.compare(part_of(data), part_of(saved(package)), repaired=True, main_part=True)
    assert set(report.counts) <= {"read-as-word", "read-through"}
    assert report.counts["read-as-word"] >= 1


@pytest.mark.parametrize("name", READ)
def test_an_untouched_save_keeps_the_bytes(name: str) -> None:
    data = source(name)
    assert part_of(saved(load(data))) == part_of(data)


def test_the_save_s_check_reads_the_source_as_the_parse_did() -> None:
    """A finding the branch held is the source's, so a save with its check on doesn't call it the session's."""
    held = '<w:document><w:body><w:p><w:pPr><w:ind w:left="99999"/></w:pPr><w:r><w:t>WORDS</w:t></w:r></w:p></w:body></w:document>'
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback>{held}</mc:Fallback></mc:AlternateContent>'
    package = load(with_part(xml.encode()))
    assert [finding.code for finding in misplaced(package)] == ["value.out_of_range"]
    package.main_document_part.contents
    assert b"WORDS" in part_of(saved(package, check=True))


def test_extract_then_save_keeps_the_bytes() -> None:
    data = source("rw-20-mc-root-choice")
    package = load(data)
    extract(package)
    assert part_of(saved(package)) == part_of(data)


# --- who asks, and what it costs -------------------------------------------------


def test_another_part_never_asks() -> None:
    assert xml_part.XmlPart.reads_root is False
    package = load(source("or-00-control"))
    assert package.main_document_part.reads_root is True


def test_a_clean_main_part_never_reaches_the_pass(monkeypatch: pytest.MonkeyPatch) -> None:
    def never(*_args: Any) -> None:
        raise AssertionError("read_root ran")

    monkeypatch.setattr(repeated, "read_root", never)
    for path in sorted(SAMPLES.rglob("*.docx"))[:6]:
        load(path.read_bytes()).main_document_part.contents


def test_two_bodies_run_the_one_repair(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[int] = []
    real = read_through.repair_bytes

    def counted(data: bytes, dropped: list[Any] | None = None) -> Any:
        calls.append(1)
        return real(data, dropped)

    monkeypatch.setattr(read_through, "repair_bytes", counted)
    load(source("rc-10-body")).main_document_part.contents
    assert calls == [1]


@pytest.mark.parametrize(
    ("data", "tag"),
    [
        (b"<a/>", "a"),
        (b'<?xml version="1.0"?><!-- c --><?pi x?><a/>', "a"),
        (b"\xef\xbb\xbf<a/>", "a"),
        ("<a/>".encode("utf-16"), "a"),
        (b"<!-- " + b"x" * 5000 + b" --><a/>", "a"),
        (f'<w:document xmlns:w="{W_NS}"/>'.encode(), f"{W}document"),
        (b"", None),
        (b"not xml", None),
        (b"<a", "a"),  # lxml gives the name, and the parse that follows refuses the part
    ],
)
def test_the_root_s_name_is_read_from_its_start_tag(data: bytes, tag: str | None) -> None:
    assert xml_part._root_tag(data) == tag


def test_the_root_s_name_doesn_t_need_the_rest_to_be_well_formed() -> None:
    assert xml_part._root_tag(b"<a><b></a>") == "a"
