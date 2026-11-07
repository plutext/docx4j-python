"""CR-040: an ``mc:AlternateContent`` that selects nothing, read as Word reads it (WD REQ-305 and REQ-306).

The oracle is the Word bridge's files (``tests/fixtures/nothing-selected/``): ``rw-25``, the ``req305``
check and the ``understood-choice`` round. Word shows the ``mc:Choice`` of a namespace it understands,
and neither branch of one it doesn't.
"""

from __future__ import annotations

import hashlib
import io
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import test_mce_levels as levels  # noqa: E402

from docx4j_py import load  # noqa: E402
from docx4j_py.child import understood_namespaces  # noqa: E402
from docx4j_py.model.text import TextOptions, TextWarning, extract, plain_text  # noqa: E402
from docx4j_py.openpackaging import LoadOptions, repeated  # noqa: E402
from docx4j_py.traversal import misplaced  # noqa: E402

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "nothing-selected"
ROOTS = Path(__file__).resolve().parents[1] / "fixtures" / "repeated"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
DOC = "/word/document.xml"

NAMES = sorted(
    path.stem for path in FIXTURES.glob("*.docx") if not path.stem.endswith("-word365") and not path.stem.startswith("te-")
)
#: Word showed neither branch of these.
NOTHING = [
    "c305-10-unknown-urn-block",
    "c305-11-unknown-urn-run",
    "c305-12-unknown-ms-block",
    "rw-25-body-choice-no-fallback",
    "uc-01-unknown-ms",
    "uc-11-cx5",
    "uc-14-cx8",
    "uc-21-mo",
    "uc-23-arto",
    "uc-25-dcmitype",
]
SELECTED = [name for name in NAMES if name not in NOTHING]
BLOCK = [name for name in NOTHING if name != "c305-11-unknown-urn-run"]


def source(name: str) -> bytes:
    return (FIXTURES / f"{name}.docx").read_bytes()


def word_text(name: str) -> str:
    text = (FIXTURES / f"{name}-word365.txt").read_text(encoding="utf-8")
    return " ".join(text.split())


def words(result: Any) -> str:
    return " ".join(plain_text(result).split())


def part_of(data: bytes) -> bytes:
    return zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml")


def saved(package: Any) -> bytes:
    out = io.BytesIO()
    package.save(out)
    return out.getvalue()


def nothing(result: Any) -> list[TextWarning]:
    return [warning for warning in result.warnings if warning.kind == TextWarning.MCE_NOTHING_SELECTED]


def with_part(xml: str) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source(NOTHING[0]))) as zin, zipfile.ZipFile(out, "w") as zout:
        for info in zin.infolist():
            zout.writestr(info, xml.encode() if info.filename == "word/document.xml" else zin.read(info.filename))
    return out.getvalue()


def document(body: str) -> str:
    return f'<w:document xmlns:w="{W_NS}" xmlns:mc="{MC_NS}" xmlns:a="urn:a" xmlns:b="urn:b"><w:body>{body}</w:body></w:document>'


CHOICE = '<mc:Choice Requires="{}"><w:p><w:r><w:t>WORDS</w:t></w:r></w:p></mc:Choice>'


def test_every_fixture_is_the_file_word_read() -> None:
    listed = dict(reversed(line.split()) for line in (FIXTURES / "SHA1SUMS").read_text().splitlines())
    assert set(listed) == {path.name for path in FIXTURES.iterdir()} - {"SHA1SUMS", "README.md"}
    for name, digest in listed.items():
        assert hashlib.sha1((FIXTURES / name).read_bytes()).hexdigest() == digest, name


@pytest.mark.parametrize("name", NAMES)
def test_the_words_are_the_words_word_shows(name: str) -> None:
    assert words(extract(load(source(name)))) == word_text(name)


@pytest.mark.parametrize("name", NOTHING)
def test_selecting_nothing_says_so(name: str) -> None:
    [warning] = nothing(extract(load(source(name))))
    assert warning.kind == "mce.nothing_selected"
    assert warning.part_name == DOC
    assert warning.selector == DOC
    assert warning.address is not None
    assert warning.detail
    assert all(uri in warning.to_dict()["detail"] for uri in warning.detail.split())


@pytest.mark.parametrize("name", SELECTED)
def test_a_choice_word_understands_says_nothing(name: str) -> None:
    assert not nothing(extract(load(source(name))))


def test_the_detail_is_the_namespaces_required_each_once_as_uris() -> None:
    body = f'<mc:AlternateContent>{CHOICE.format("a b")}{CHOICE.format("b")}</mc:AlternateContent>'
    [warning] = nothing(extract(load(with_part(document(body)))))
    assert warning.detail == "urn:a urn:b"


def test_each_element_is_its_own_warning() -> None:
    element = f'<mc:AlternateContent>{CHOICE.format("a")}</mc:AlternateContent>'
    found = nothing(extract(load(with_part(document(element + element)))))
    assert [warning.address for warning in found] == ["body/0", "body/1"]


def test_an_element_with_no_branch_says_nothing() -> None:
    assert not nothing(extract(load(with_part(document("<mc:AlternateContent/>")))))


def test_a_fallback_is_still_selected() -> None:
    fallback = "<mc:Fallback><w:p><w:r><w:t>FALLBACK</w:t></w:r></w:p></mc:Fallback>"
    body = f'<mc:AlternateContent>{CHOICE.format("a")}{fallback}</mc:AlternateContent>'
    result = extract(load(with_part(document(body))))
    assert words(result) == "FALLBACK"
    assert not nothing(result)


def test_the_warnings_option_turns_it_off() -> None:
    assert not nothing(extract(load(source(NOTHING[0])), TextOptions(warnings=False)))


@pytest.mark.parametrize("name", NOTHING)
def test_word_mode_reads_the_same_and_says_the_same(name: str) -> None:
    found = extract(load(source(name), options=LoadOptions(mc_load_mode="word")))
    assert words(found) == word_text(name)
    assert [(w.part_name, w.detail) for w in nothing(found)] == [(w.part_name, w.detail) for w in nothing(extract(load(source(name))))]


@pytest.mark.parametrize("name", BLOCK)
def test_word_mode_takes_the_element_out_as_word_s_save_does(name: str) -> None:
    package = load(source(name), options=LoadOptions(mc_load_mode="word"))
    (record,) = _read(package).main_document_part.resolved_alternate_content
    assert record.branch is None
    assert b"AlternateContent" not in part_of(saved(package))


def _read(package: Any) -> Any:
    package.main_document_part.contents
    return package


ELEMENT = f'<mc:AlternateContent>{CHOICE.format("a")}</mc:AlternateContent>'
GRID = '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
PLACES = {
    "only child of the body": ELEMENT,
    "last in the body": f"<w:p/>{ELEMENT}",
    "two in a row": f"<w:p/>{ELEMENT}{ELEMENT}<w:p/>",
    "two at the end": f"<w:p/>{ELEMENT}{ELEMENT}",
    "only child of a cell": f"<w:tbl>{GRID}<w:tr><w:tc>{ELEMENT}</w:tc></w:tr></w:tbl><w:p/>",
    "last in a cell": f"<w:tbl>{GRID}<w:tr><w:tc><w:p/>{ELEMENT}</w:tc></w:tr></w:tbl><w:p/>",
    "only child of a row": f'<w:tbl>{GRID}<w:tr><mc:AlternateContent><mc:Choice Requires="a"><w:tc><w:p/></w:tc></mc:Choice></mc:AlternateContent></w:tr></w:tbl><w:p/>',
    "only child of a table": f'<w:tbl>{GRID}<mc:AlternateContent><mc:Choice Requires="a"><w:tr><w:tc><w:p/></w:tc></w:tr></mc:Choice></mc:AlternateContent></w:tbl><w:p/>',
}


@pytest.mark.parametrize("name", ["te-01-tc-only", "te-02-tc-only-one-cell", "te-03-tc-no-branch"])
def test_a_cell_left_with_no_block_keeps_a_paragraph_as_word_s_repair_gives_it(name: str) -> None:
    """A cell with no block makes Word offer a repair, so word mode writes what the repair writes."""
    package = load(source(name), options=LoadOptions(mc_load_mode="word"))
    root = etree.fromstring(part_of(saved(_read(package))))

    def cells(tree: Any) -> list[list[str]]:
        return [["".join(p.itertext()) for p in cell.findall(f"{W}p")] for cell in tree.iter(f"{W}tc")]

    theirs = etree.fromstring(part_of((FIXTURES / f"{name}-word365.docx").read_bytes()))
    assert cells(root) == cells(theirs)
    assert not [finding for finding in misplaced(package) if finding.code == "cell.no_end_paragraph"]


def test_a_cell_keeps_only_the_paragraph_it_needs() -> None:
    held = f"<w:tbl>{GRID}<w:tr><w:tc>{ELEMENT}{ELEMENT}<w:p/></w:tc></w:tr></w:tbl><w:p/>"
    package = load(with_part(document(held)), options=LoadOptions(mc_load_mode="word"))
    root = etree.fromstring(part_of(saved(_read(package))))
    assert len(next(root.iter(f"{W}tc")).findall(f"{W}p")) == 1


def test_a_cell_emptied_twice_gets_one_paragraph() -> None:
    held = f"<w:tbl>{GRID}<w:tr><w:tc>{ELEMENT}{ELEMENT}</w:tc></w:tr></w:tbl><w:p/>"
    package = load(with_part(document(held)), options=LoadOptions(mc_load_mode="word"))
    root = etree.fromstring(part_of(saved(_read(package))))
    assert len(next(root.iter(f"{W}tc")).findall(f"{W}p")) == 1


@pytest.mark.parametrize("place", PLACES)
def test_word_mode_says_what_preserve_says_wherever_the_element_stood(place: str) -> None:
    """Taken out at a container's end, or as all it held, no item stands after it for the walk to meet."""
    data = with_part(document(PLACES[place]))
    kept = nothing(extract(load(data)))
    taken = nothing(extract(load(data, options=LoadOptions(mc_load_mode="word"))))
    assert kept
    assert [(w.address, w.part_name, w.detail) for w in taken] == [(w.address, w.part_name, w.detail) for w in kept]


@pytest.mark.parametrize("name", NOTHING)
def test_a_default_load_keeps_the_element_in_a_save(name: str) -> None:
    assert b"AlternateContent" in part_of(saved(_read(load(source(name)))))


@pytest.mark.parametrize("name", NOTHING)
def test_mce_preprocess_reads_the_same_and_says_nothing(name: str) -> None:
    result = extract(load(source(name), options=LoadOptions(mce_preprocess=True)))
    assert words(result) == word_text(name)
    assert not nothing(result)


def test_the_selector_is_the_part_it_stands_in() -> None:
    """A note, not the main part: the warning names the story's part."""
    _kind, made, _part_of, _content_of = levels.CASES["w:footnote"]
    probe = made()
    for name, xml in probe.parts.items():
        if isinstance(xml, str) and "AlternateContent" in xml:
            xml = re.sub(r'Requires="w14"', 'Requires="w98"', xml)
            probe.parts[name] = re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", xml, flags=re.S)
    [warning] = nothing(extract(probe.load("preserve")))
    assert warning.part_name == "/word/footnotes.xml"


# --- at a main part's root -------------------------------------------------------


def test_a_root_that_selects_nothing_says_so() -> None:
    package = load((ROOTS / "rw-24-mc-root-nothing-selected.docx").read_bytes())
    result = extract(package)
    assert words(result) == ""
    [warning] = nothing(result)
    assert (warning.address, warning.part_name, warning.detail) == (None, DOC, "urn:docx4j:root")
    [record] = package.main_document_part.warnings
    assert isinstance(record, repeated.ReadAsWord)
    assert record.nothing_selected == "urn:docx4j:root"
    assert record.to_dict()["nothing_selected"] == "urn:docx4j:root"


def test_a_root_with_a_branch_says_nothing_of_it() -> None:
    assert not nothing(extract(load((ROOTS / "rw-20-mc-root-choice.docx").read_bytes())))


def test_a_root_with_no_branch_says_nothing_of_it() -> None:
    package = load(with_part(f'<mc:AlternateContent xmlns:mc="{MC_NS}"/>'))
    assert not nothing(extract(package))


def test_a_prefix_nothing_binds_gives_no_uri() -> None:
    package = load(with_part(f'<mc:AlternateContent xmlns:mc="{MC_NS}"><mc:Choice Requires="unbound"/></mc:AlternateContent>'))
    assert [warning.detail for warning in nothing(extract(package))] == [""]


def test_a_root_s_nothing_selected_under_a_branch_says_so() -> None:
    inner = '<mc:AlternateContent><mc:Choice Requires="z"/></mc:AlternateContent>'
    xml = f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:z="urn:z"><mc:Fallback>{inner}</mc:Fallback></mc:AlternateContent>'
    assert [warning.detail for warning in nothing(extract(load(with_part(xml))))] == ["urn:z"]


# --- the two sets ----------------------------------------------------------------


@pytest.mark.parametrize(
    "uri",
    [
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "http://schemas.microsoft.com/office/2019/extlst",
        "http://schemas.microsoft.com/office/2006/activeX",
        "http://www.w3.org/2001/XMLSchema-instance",
    ],
)
def test_the_understood_set_holds_what_word_was_measured_to_understand(uri: str) -> None:
    assert uri in understood_namespaces()


@pytest.mark.parametrize(
    "uri",
    [
        "http://purl.org/dc/elements/1.1/",
        "http://schemas.microsoft.com/office/2006/encryption",
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        "urn:schemas-microsoft-com:office:excel",
    ],
)
def test_the_understood_set_leaves_out_what_word_was_measured_not_to(uri: str) -> None:
    assert uri not in understood_namespaces()
