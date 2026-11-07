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
    "rq-10-madeup-ac",
    "rq-30-branch-sdt",
    "rq-32-branch-document-then-hdr",
    "rq-33-branch-hdr-then-document",
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
    "rq-11-madeup-sdt-runs",
    "rq-21-sdtcontent-rows",
    "rq-22-sdtcontent-row-sdt",
    "rq-31-branch-txbxcontent",
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


@pytest.mark.parametrize("name", ["rc-00-control", "rw-00-control", "rw-01-control-mc-choice", "or-00-control", "rq-00-control"])
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
        """Each paragraph's words, but a block-level ``mc:Fallback``'s: preserve keeps every branch (CR-012), ``rq-10``'s."""
        body = etree.fromstring(xml).find(f"{W}body")
        fallback = f"{{{MC_NS}}}Fallback"
        return [
            text
            for paragraph in body.iter(f"{W}p")
            if not any(up.tag == fallback for up in paragraph.iterancestors())
            if (text := "".join(paragraph.itertext()))
        ]

    assert words(ours) == words(theirs)
    if page_widths(ours):  # a file that states no section has none, where Word's save writes a default
        assert page_widths(ours) == page_widths(theirs)


def whole_sections(xml: bytes) -> list[tuple[dict[str, str], dict[str, str]]]:
    """Each section's ``w:pgSz`` and ``w:pgMar``, every attribute: a width alone can't see a surviving ``w:orient``."""
    out = []
    for section in etree.fromstring(xml).iter(f"{W}sectPr"):
        size, margins = section.find(f"{W}pgSz"), section.find(f"{W}pgMar")
        out.append(tuple({etree.QName(k).localname: v for k, v in (e.attrib.items() if e is not None else ())} for e in (size, margins)))
    return out


@pytest.mark.parametrize("name", [f"nd-2{i}-{rest}" for i, rest in enumerate(
    ["body-only-child", "body-middle-no-sectPr", "body-middle-with-sectPr", "body-last-with-sectPr", "body-three-deep"])])
def test_a_body_in_a_body_keeps_word_s_whole_section(name: str) -> None:
    """``nd-22``, ``nd-24``: the nested body's section and the later one's, read as one (REQ-322), as Word's save holds
    them: the later size and margins, the nested one's ``w:orient="landscape"`` outliving them (docx4j's correction,
    2026-11-08)."""
    package = load(source(name))
    package.main_document_part.contents
    assert whole_sections(part_of(saved(package))) == whole_sections(word_save(name))


def test_an_unknown_root_holding_mc_alternate_content_saves_as_word_s_under_word_mode() -> None:
    """``rq-10``: the branch read where it stands; ``mc_load_mode="word"`` writes it as Word's save does."""
    package = load(source("rq-10-madeup-ac"), options=LoadOptions(mc_load_mode="word"))
    assert plain_text(extract(package)).strip() == word_text("rq-10-madeup-ac") == "CHOICE WORDS"
    package.main_document_part.contents  # a reading is private; this marks the part for writing
    ours = etree.fromstring(part_of(saved(package)))
    assert ["".join(p.itertext()) for p in ours.iter(f"{W}p")] == ["CHOICE WORDS"]
    assert ours.find(f".//{{{MC_NS}}}AlternateContent") is None


def test_word_mode_takes_a_branch_an_unknown_element_held_in_a_document_too() -> None:
    """The same order of passes as ``rq-10``'s, below a ``w:document`` root: read through, then spliced."""
    branch = (
        f'<mc:AlternateContent xmlns:mc="{MC_NS}"><mc:Choice Requires="w14"><w:p><w:r><w:t>CHOICE</w:t></w:r></w:p></mc:Choice>'
        "<mc:Fallback><w:p><w:r><w:t>FALLBACK</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>"
    )
    xml = f'<w:document xmlns:w="{W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:madeUp>{branch}</w:madeUp></w:body></w:document>'
    package = load(with_part(xml.encode()), options=LoadOptions(mc_load_mode="word"))
    assert plain_text(extract(package)).strip() == "CHOICE"
    package.main_document_part.contents  # a reading is private; this marks the part for writing
    assert len(package.main_document_part.resolved_alternate_content) == 1
    assert b"FALLBACK" not in part_of(saved(package)) and b"madeUp" not in part_of(saved(package))
    kept = load(with_part(xml.encode()))
    kept.main_document_part.contents
    assert plain_text(extract(kept)).strip() == "CHOICE"
    assert b"FALLBACK" in part_of(saved(kept)) and b"madeUp" not in part_of(saved(kept))


def test_a_control_s_level_is_what_its_content_holds() -> None:
    """A run-level control in an unknown root is refused (``rq-11``); one in a ``w:sdtContent`` root that holds a row too
    (``rq-22``); a block one as the root is read (``rq-30``). A run-level one as the root is unmeasured: REQ-320's."""
    sdt = '<w:sdt><w:sdtContent><w:r><w:t>RUN</w:t></w:r></w:sdtContent></w:sdt>'
    package = load(with_part(f'<w:sdt xmlns:w="{W_NS}"><w:sdtContent><w:r><w:t>RUN</w:t></w:r></w:sdtContent></w:sdt>'.encode()))
    assert [(w.kind, w.detail) for w in extract(package).warnings] == [("part.unexpected_root", "w:sdt")]
    block = load(with_part(f'<w:sdt xmlns:w="{W_NS}"><w:sdtContent><w:p><w:r><w:t>BLOCK</w:t></w:r></w:p></w:sdtContent></w:sdt>'.encode()))
    assert plain_text(extract(block)).strip() == "BLOCK"
    with pytest.raises(Docx4JException):
        load(with_part(f'<w:madeUp xmlns:w="{W_NS}"><w:p/>{sdt}</w:madeUp>'.encode())).main_document_part.contents


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


def test_two_bodies_keep_one_section_and_list_what_it_replaced() -> None:
    package = load(source("rc-10-body"))
    part = package.main_document_part
    body = part.contents.body
    assert len([item for item in body.content]) == 2
    assert body.sect_pr is not None
    assert local_names(part.skipped) == ["pgSz", "pgMar"]
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


def test_a_body_with_no_section_is_passed_over() -> None:
    """``rt-17``: the third body has no ``w:sectPr``, so the second's is the last applied, and what it replaced is listed."""
    name = "rt-17-body-third-no-sectPr"
    package = load(source(name))
    assert page_widths(part_of(saved(_read(package)))) == page_widths(word_save(name)) == ["15840"]
    assert local_names(package.main_document_part.skipped) == ["pgSz", "pgMar"]


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


def test_a_later_section_is_applied_in_one_body() -> None:
    name = "rc-40-sectPr"
    ours = part_of(saved(_read(load(source(name)))))
    assert page_widths(ours) == page_widths(word_save(name)) == ["15840"]


def test_later_table_properties_are_applied() -> None:
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


def test_a_run_level_control_joins_too() -> None:
    root, read_as_one = repaired("<w:p><w:sdt><w:sdtContent/><w:sdtContent/></w:sdt></w:p>")
    assert read_as_one == 1
    assert len(list(root.iter(f"{W}sdtContent"))) == 1


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


@pytest.mark.parametrize(
    "name",
    [
        "jc", "b", "sz", "pPr", "rPr", "tblPr", "sectPr", "sdtContent", "body",
        "ind", "spacing", "rFonts", "pgSz", "pgMar", "cols", "vanish", "numPr",
    ],
)
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


# --- Word's merge rule, REQ-322 ----------------------------------------------------

#: Word's own defaults, written where a file states nothing: no reading of it.
WORD_ADDS = {"tblPr": {"tblCellMar", "tblLook"}}
#: REQ-322 says nothing of ``w:tblW``: ``w:w="4000"`` then ``w:type="auto"`` saves as ``w:w="0" w:type="dxa"``, one file.
MERGED = sorted(path.stem for path in FIXTURES.glob("sm-*.docx") if not path.stem.endswith("-word365") and "-61-" not in path.stem)


def _stated(element: etree._Element) -> dict[str, dict[str, str]]:
    """`element`'s children by name, each with its attributes, as a save says them whoever wrote it."""
    facts: dict[str, dict[str, str]] = {}
    for child in element:
        name = etree.QName(child).localname
        if name in WORD_ADDS.get(etree.QName(element).localname, ()) or name == "headerReference":
            continue
        said = {etree.QName(key).localname: value for key, value in child.attrib.items()}
        said = {key: value for key, value in said.items() if not key.startswith("rsid")}
        if said.get("orient") == "portrait":
            del said["orient"]  # the default, which Word's save leaves out (sm-07)
        if name == "cols":
            said.setdefault("space", "720")  # likewise the other way (sm-32)
        facts[name] = said
    if etree.QName(element).localname == "sectPr":
        facts.setdefault("cols", {"space": "720"})
        facts["headerReference"] = {"count": str(len(element.findall(f"{W}headerReference")))}
    return facts


def _merged(xml: bytes) -> dict[str, Any]:
    """What the ``sm-*`` round asks of a save: every section, every table's properties, the target's properties."""
    root = etree.fromstring(xml)
    target = next(p for p in root.iter(f"{W}p") if "TARGET TEXT" in "".join(p.itertext()))
    runs = [properties for run in target.findall(f"{W}r") for properties in run.findall(f"{W}rPr")]
    return {
        "sections": [_stated(section) for section in root.iter(f"{W}sectPr")],
        "tables": [_stated(properties) for properties in root.iter(f"{W}tblPr")],
        "paragraph": [_stated(properties) for properties in target.findall(f"{W}pPr") if len(properties)],
        "runs": [_stated(properties) for properties in runs if len(properties)],
    }


def _hidden_off(facts: dict[str, Any]) -> dict[str, Any]:
    """``sm-80``: Word's save of off writes no ``w:vanish``, and here it's the last as written, which reads the same."""
    facts["runs"] = [run for run in facts["runs"] if run != {"vanish": {"val": "0"}}]
    return facts


@pytest.mark.parametrize("name", MERGED)
def test_a_save_holds_what_word_s_save_holds(name: str) -> None:
    ours = part_of(saved(_read(load(source(name)))))
    assert _hidden_off(_merged(ours)) == _merged(word_save(name))


@pytest.mark.parametrize("name", MERGED)
def test_a_merge_is_all_a_save_changes(name: str) -> None:
    report = canon.compare(part_of(source(name)), part_of(saved(_read(load(source(name))))), repaired=True, main_part=True)
    assert report.identical, report.diffs[:5]


def _properties(xml: str) -> dict[str, dict[str, str]]:
    """The first paragraph's ``w:pPr`` as a save holds it, of a body holding `xml`."""
    data = f'<w:document xmlns:w="{W_NS}"><w:body>{xml}</w:body></w:document>'.encode()
    root = etree.fromstring(part_of(saved(_read(load(with_part(data))))))
    return _stated(next(root.iter(f"{W}pPr")))


def test_copies_apply_in_order_each_holder_s_own_first() -> None:
    """Holders in order and copies in a holder in order, so the third ``w:ind`` wins over the second."""
    xml = (
        '<w:p><w:pPr><w:ind w:left="1"/><w:ind w:left="2" w:right="2"/></w:pPr><w:pPr><w:ind w:left="3"/></w:pPr>'
        "<w:r><w:t>A</w:t></w:r></w:p>"
    )
    assert _properties(xml)["ind"] == {"left": "3", "right": "2"}


def test_a_copy_that_states_the_same_loses_nothing() -> None:
    xml = (
        '<w:document xmlns:w="{}"><w:body><w:p><w:pPr><w:jc w:val="left"/><w:ind w:left="1"/></w:pPr><w:pPr><w:jc w:val="left"/>'
        '<w:ind w:left="1"/></w:pPr></w:p></w:body></w:document>'
    ).format(W_NS)
    package = _read(load(with_part(xml.encode())))
    assert not package.main_document_part.skipped


def test_an_attribute_replaced_is_listed_and_one_added_isn_t() -> None:
    assert local_names(_read(load(source("sm-02-pgsz-orient-last"))).main_document_part.skipped) == ["pgSz"]
    assert not _read(load(source("sm-71-ind-across-ppr"))).main_document_part.skipped


@pytest.mark.parametrize(("first", "last", "on"), [('w:val="0"', "", True), ("", 'w:val="0"', False)])
def test_the_last_on_off_element_decides_across_holders_too(first: str, last: str, on: bool) -> None:
    """An element with no ``w:val`` states on, so no attribute of an earlier one stays (``sm-81``)."""
    run = f"<w:r><w:rPr><w:vanish {first}/></w:rPr><w:rPr><w:vanish {last}/></w:rPr><w:t>A</w:t></w:r>"
    data = f'<w:document xmlns:w="{W_NS}"><w:body><w:p>{run}</w:p></w:body></w:document>'.encode()
    [vanish] = etree.fromstring(part_of(saved(_read(load(with_part(data)))))).iter(f"{W}vanish")
    assert (vanish.get(f"{W}val") is None) is on


def test_a_child_the_table_doesn_t_name_is_the_later_s_whole() -> None:
    """Unmeasured, so ``w:shd`` doesn't merge: the later stands and the earlier is listed."""
    xml = (
        f'<w:document xmlns:w="{W_NS}"><w:body><w:p><w:pPr><w:shd w:val="clear" w:fill="FF0000"/></w:pPr>'
        '<w:pPr><w:shd w:val="clear"/></w:pPr></w:p></w:body></w:document>'
    )
    package = _read(load(with_part(xml.encode())))
    root = etree.fromstring(part_of(saved(package)))
    assert _stated(next(root.iter(f"{W}pPr")))["shd"] == {"val": "clear"}
    assert local_names(package.main_document_part.skipped) == ["shd"]


@pytest.mark.parametrize(("name", "words"), [("sm-40-header-first", "HEADER ONE"), ("sm-42-header-both", "HEADER ONE"), ("sm-41-header-last", "")])
def test_only_the_first_copy_s_header_stands(name: str, words: str) -> None:
    """A later copy's reference is ignored, with its header's words, and listed."""
    package = load(source(name))
    text = plain_text(extract(package, include=["headers"]))
    assert ("HEADER" in text) is bool(words) and "HEADER TWO" not in text
    assert ("headerReference" in local_names(package.main_document_part.skipped)) is (name != "sm-40-header-first")


def test_a_later_footer_reference_is_listed() -> None:
    """Word's ``ft-*``: the first copy's stands, and what a later one named is listed."""
    def section(rid: str) -> str:
        return f'<w:sectPr><w:footerReference w:type="default" r:id="{rid}"/></w:sectPr>'

    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xml = f'<w:document xmlns:w="{W_NS}" xmlns:r="{r_ns}"><w:body><w:p/>{section("rId1")}{section("rId2")}</w:body></w:document>'
    package = _read(load(with_part(xml.encode())))
    root = etree.fromstring(part_of(saved(package, check=False)))
    assert [ref.get(f"{{{r_ns}}}id") for ref in root.iter(f"{W}footerReference")] == ["rId1"]
    assert local_names(package.main_document_part.skipped) == ["footerReference"]


def _references(data: bytes) -> list[list[tuple[str, str, str]]]:
    """Each section's header and footer references: their kind, type and part's words, as a save's relationships say."""
    package = zipfile.ZipFile(io.BytesIO(data))
    targets = {rel.get("Id"): rel.get("Target") for rel in etree.fromstring(package.read("word/_rels/document.xml.rels"))}
    r_id = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    found = []
    for section in etree.fromstring(package.read("word/document.xml")).iter(f"{W}sectPr"):
        held = [ref for ref in section if etree.QName(ref).localname in ("headerReference", "footerReference")]
        words = ["".join(etree.fromstring(package.read(f"word/{targets[ref.get(r_id)]}")).itertext()).strip() for ref in held]
        found.append(sorted((etree.QName(ref).localname, ref.get(f"{W}type"), said) for ref, said in zip(held, words, strict=True)))
    return found


FOOTERS = sorted(path.stem for path in FIXTURES.glob("ft-*.docx") if not path.stem.endswith("-word365"))


@pytest.mark.parametrize("name", FOOTERS)
def test_only_the_first_copy_s_references_stand(name: str) -> None:
    """Of any ``w:type``, across joined bodies too, and a first copy with none leaves none."""
    word = (FIXTURES / f"{name}-word365.docx").read_bytes()
    assert _references(saved(_read(load(source(name))))) == _references(word)


# --- WD 0.101: w:numPr, w:sdtContent at every level, an mc root in other parts ---------


def test_a_num_pr_is_applied_to_one_object() -> None:
    """``c79-41``: the first's ``w:ilvl`` with the second's ``w:numId``, and the label that gives."""
    name = "c79-41-ppr-numpr"
    package = _read(load(source(name)))
    ours, theirs = (
        [{etree.QName(child).localname: child.get(f"{W}val") for child in held} for held in etree.fromstring(xml).iter(f"{W}numPr")]
        for xml in (part_of(saved(package)), word_save(name))
    )
    assert ours == theirs == [{"ilvl": "0", "numId": "2"}]
    assert plain_text(extract(load(source(name)), list_labels=True)).split() == word_text(name).split()


@pytest.mark.parametrize(
    ("name", "words"),
    [
        ("c79-30-sdtcontent-run", ["FIRST RUNSECOND RUN"]),
        ("c79-31-sdtcontent-row", ["FIRST ROW", "SECOND ROW", "AFTER"]),
        ("c79-32-sdtcontent-cell", ["FIRST CELL", "SECOND CELL", "AFTER"]),
    ],
)
def test_sdt_content_joins_at_every_level(name: str, words: list[str]) -> None:
    package = load(source(name))
    text = plain_text(extract(package))
    assert [line for part in text.split("\n") for line in part.split("\t") if line] == words
    assert [line for line in word_text(name).replace("\r", "\n").split("\n") if line.strip()] == words
    assert not package.main_document_part.skipped


@pytest.mark.parametrize("name", ["c79-40-ppr-nested-rpr", "c79-42-ppr-pprchange"])
def test_what_an_earlier_copy_alone_states_stays(name: str) -> None:
    def named(xml: bytes) -> list[list[str]]:
        return [[etree.QName(child).localname for child in held] for held in etree.fromstring(xml).iter(f"{W}pPr")][:1]

    assert named(part_of(saved(_read(load(source(name)))))) == named(word_save(name))


#: ``c79-10`` to ``c79-14`` and ``c79-18``: the words Word's branch holds, and what's asked to show them.
ROOTS = {
    "c79-10-hdr-root": ("CHOICE HEADER", {"include": ["headers"]}),
    "c79-11-ftr-root": ("CHOICE FOOTER", {"include": ["footers"]}),
    "c79-12-footnotes-root": ("CHOICE NOTE", {}),
    "c79-13-endnotes-root": ("CHOICE NOTE", {}),
    "c79-14-comments-root": ("CHOICE COMMENT", {"comments": True}),
    "c79-18-hdr-root-unknown-choice": ("FALLBACK HEADER", {"include": ["headers"]}),
}


@pytest.mark.parametrize(("name", "held"), ROOTS.items(), ids=ROOTS)
def test_an_mc_root_in_another_part_is_the_branch_word_takes(name: str, held: tuple[str, dict[str, Any]]) -> None:
    words, options = held
    assert words in word_text(name)
    text = plain_text(extract(load(source(name)), **options))
    assert words in text and "FALLBACK" not in text.replace(words, "")


def test_a_numbering_root_reaches_the_emulator() -> None:
    """It reads the part's bytes itself, so it's the case ``XmlPart.read_xml`` is for."""
    name = "c79-16-numbering-root"
    assert plain_text(extract(load(source(name)), list_labels=True)).split() == word_text(name).split()


def test_a_styles_root_is_read() -> None:
    package = load(source("c79-15-styles-root"))
    assert "Probe" in {style.style_id for style in package.main_document_part.style_definitions_part.contents.style}


def test_a_settings_root_is_read() -> None:
    name = "c79-17-settings-root"
    package = load(source(name))
    theirs = zipfile.ZipFile(FIXTURES / f"{name}-word365.docx").read("word/settings.xml")
    tab = etree.fromstring(theirs).find(f"{W}defaultTabStop").get(f"{W}val")
    assert str(package.main_document_part.document_settings_part.contents.default_tab_stop.val) == tab == "1000"


def test_an_mc_root_read_is_said_and_what_went_listed() -> None:
    package = load(source("c79-10-hdr-root"))
    [part] = package.main_document_part.header_parts()
    part.contents
    assert [(w.root, w.children) for w in part.warnings] == [("mc:AlternateContent", 0)]
    assert local_names(part.skipped) == ["Fallback"]


def test_an_untouched_mc_root_is_saved_byte_for_byte() -> None:
    package = load(source("c79-10-hdr-root"))
    plain_text(extract(package, include=["headers"]))
    name = "word/header1.xml"
    assert zipfile.ZipFile(io.BytesIO(saved(package))).read(name) == zipfile.ZipFile(io.BytesIO(source("c79-10-hdr-root"))).read(name)


def test_a_header_read_for_real_is_saved_as_its_branch() -> None:
    package = load(source("c79-10-hdr-root"))
    [part] = package.main_document_part.header_parts()
    part.contents
    root = etree.fromstring(zipfile.ZipFile(io.BytesIO(saved(package))).read("word/header1.xml"))
    assert root.tag == f"{W}hdr" and "".join(root.itertext()) == "CHOICE HEADER"


def _header_root(held: str) -> bytes:
    """``c79-10`` with its header's root `held`."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source("c79-10-hdr-root"))) as zin, zipfile.ZipFile(out, "w") as zout:
        for info in zin.infolist():
            zout.writestr(info, held.encode() if info.filename == "word/header1.xml" else zin.read(info.filename))
    return out.getvalue()


def test_a_part_selecting_nothing_is_empty_and_says_so() -> None:
    xml = (
        f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}" xmlns:x="urn:not-understood">'
        '<mc:Choice Requires="x"><w:hdr><w:p><w:r><w:t>GONE</w:t></w:r></w:p></w:hdr></mc:Choice></mc:AlternateContent>'
    )
    result = extract(load(_header_root(xml)), include=["headers"])
    assert "GONE" not in plain_text(result)
    assert [(w.kind, w.selector, w.detail) for w in result.warnings] == [("mce.nothing_selected", "/word/header1.xml", "urn:not-understood")]


def test_a_branch_holding_something_else_stays_as_it_was() -> None:
    """Unmeasured, so nothing is read, and the parse lists the root's children as it always has."""
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>'
    package = load(_header_root(xml))
    [part] = package.main_document_part.header_parts()
    part.contents
    assert not part.warnings and local_names(part.skipped) == ["Fallback"]


def test_a_nested_mc_root_is_read_through_each_level_in_another_part() -> None:
    held = '<w:hdr><w:p><w:r><w:t>DEEP</w:t></w:r></w:p></w:hdr>'
    inner = f"<mc:AlternateContent><mc:Fallback>{held}</mc:Fallback></mc:AlternateContent>"
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:w="{W_NS}"><mc:Fallback>{inner}</mc:Fallback></mc:AlternateContent>'
    assert "DEEP" in plain_text(extract(load(_header_root(xml)), include=["headers"]))


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
