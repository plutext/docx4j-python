"""A ``w:br`` outside a ``w:r``, read as Word reads it: CR-030, the port of docx4j CR-073.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_br_outside_run.py -q

The oracle is docx4j's ``br-placement`` probe at ``95c6b42d3``, copied unedited
into ``tests/fixtures/br-placement/`` (its ``README.md`` has the provenance):
thirteen small files, each one shape of a ``w:br`` outside a run, and Word 365's
save of each (``*-word365.docx``). docx4j's ``BrOutsideARunTest`` compares the
two, and so does this file, with its comparison (:func:`tokens`): the text, the
breaks and their types, the field characters and the paragraph ends, in order,
each with its run's bold, italic and size, over every story part.

What is held (CR-030 section 5):

* each probe reads as Word's save of it, by that comparison and in plain text;
  ``br2-3`` with its bare ``w:instrText`` read as ``w:t`` (Q6, docx4j's D6),
  and its header's plain text pinned as that difference;
* the records: each part's ``br.outside_run`` warning, its counts and message,
  ``package.warnings``, the log, nothing left in ``part.skipped``, and no
  warning on any of Word's saves;
* ``strict=True`` loads all thirteen (Q3), a checked save reads back the same
  with no warning, ``canon.compare(..., repaired=True)`` is identical, and an
  untouched save is byte for byte;
* docx4j's edges (``BrOutsideARunTest``'s five methods) and section 2.6's own;
* the routes: ``source_tree()``, adoption, ``set_bytes``, ``insert_ooxml``'s
  fragment route (Q4) and ``wml()`` refusing;
* the trigger reads the report's path (section 2.3), and the cost of a clean
  part, deterministically: Symbols' main part is parsed once and the pass
  never runs.
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

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon

from docx4j_py import create_package, load
from docx4j_py.child import iter_children
from docx4j_py.fragments import FragmentError, wml
from docx4j_py.model.text import extract, plain_text
from docx4j_py.openpackaging import LoadOptions, XmlPart, breaks, read_through
from docx4j_py.openpackaging.breaks import BrOutsideRun, triggers
from docx4j_py.openpackaging.read_through import repair_bytes
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.parts.wml import FooterPart, HeaderPart
from docx4j_py.traversal import text_of

FIXTURES = ROOT / "tests" / "fixtures" / "br-placement"
SAMPLES = ROOT / "samples"

#: The thirteen probes, as docx4j names them.
PROBES = sorted(path.stem for path in FIXTURES.glob("*.docx") if not path.stem.endswith("-word365"))
#: The one probe whose reading differs from Word's save by Q6 (docx4j's D6) alone.
BARE_INSTR_TEXT = "br2-3-instrtext-bare-header"
#: The rest, which read exactly as Word's saves.
SAME_AS_WORD = [stem for stem in PROBES if stem != BARE_INSTR_TEXT]

#: Each probe's warnings: part, breaks, text elements (CR-030 section 2.6; docx4j's WARN).
WARNINGS: dict[str, list[tuple[str, int, int]]] = {
    "br-1-p-alone": [("/word/document.xml", 1, 0)],
    "br-2-p-between-runs": [("/word/document.xml", 1, 0)],
    "br-3-t-mixed": [("/word/document.xml", 0, 1)],
    "br-4-t-nested": [("/word/document.xml", 0, 1)],
    "br-5-hyperlink": [("/word/document.xml", 1, 0)],
    "br-6-cell": [("/word/document.xml", 1, 0)],
    "br-7-footer": [("/word/footer1.xml", 1, 0)],
    "br-8-p-page": [("/word/document.xml", 1, 0)],
    "br2-1-run-properties": [("/word/document.xml", 8, 1)],
    "br2-2-instrtext-field": [("/word/document.xml", 0, 1)],
    "br2-3-instrtext-bare-header": [("/word/header1.xml", 0, 1)],
    "br2-4-body-br": [("/word/document.xml", 1, 0)],
    "br2-5-body-br-page": [("/word/document.xml", 1, 0)],
}

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TEXTS = {W + "t": "t", W + "instrText": "instrText", W + "delText": "delText", W + "delInstrText": "delInstrText"}
PROPERTIES = {
    W + n for n in ("pPr", "rPr", "sectPr", "tblPr", "tblGrid", "trPr", "tcPr", "sdtPr", "sdtEndPr", "tblPrEx")
}

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
)


# ---------------------------------------------------------------------------
# helpers: docx4j BrOutsideARunTest's comparison
# ---------------------------------------------------------------------------


def source(stem: str) -> Path:
    return FIXTURES / f"{stem}.docx"


def word(stem: str) -> Path:
    return FIXTURES / f"{stem}-word365.docx"


def story_parts(pkg: Any) -> list[Any]:
    """The main part, then each header and footer, as ``BrOutsideARunTest`` takes them."""
    return [pkg.main_document_part] + [p for p in pkg.parts.parts() if isinstance(p, (HeaderPart, FooterPart))]


def _on(value: Any) -> bool:
    return value is not None and getattr(value, "val", None) not in (False, "0", "false", "off")


def props(r_pr: Any) -> str:
    """The properties the probes use: bold, italic, size."""
    if r_pr is None:
        return ""
    bits = []
    if _on(getattr(r_pr, "b", None)):
        bits.append("b")
    if _on(getattr(r_pr, "i", None)):
        bits.append("i")
    sz = getattr(r_pr, "sz", None)
    if sz is not None:
        bits.append(f"sz{sz.val}")
    return f"[{' '.join(bits)}]" if bits else ""


def walk(obj: Any, out: list[str]) -> None:
    """Text (merged where Word and this package split runs differently), breaks, field characters, paragraph ends."""
    for qname, child in iter_children(obj):
        if qname in PROPERTIES:
            continue
        if qname == W + "r":
            pr = props(getattr(child, "r_pr", None))
            for name, item in iter_children(child):
                if name in TEXTS:
                    head = TEXTS[name] + pr + ":"
                    text = item.value or ""
                    if out and out[-1].startswith(head):
                        out[-1] += text
                    else:
                        out.append(head + text)
                elif name == W + "br":
                    kind = getattr(item, "type_value", None)
                    out.append("br" + (f"({kind.value})" if kind is not None else "") + pr)
                elif name == W + "fldChar":
                    out.append("fld:" + item.fld_char_type.value)
            continue
        walk(child, out)
        if qname == W + "p":
            out.append("¶")


def tokens(pkg: Any) -> list[str]:
    """The main part's tokens, then each header's and footer's."""
    out: list[str] = []
    for part in story_parts(pkg):
        if isinstance(part, (HeaderPart, FooterPart)):
            out.append("header" if isinstance(part, HeaderPart) else "footer")
            walk(part.contents, out)
        else:
            walk(part.contents.body, out)
    return out


def body_tokens(body_xml: str, **options: Any) -> tuple[list[str], Any]:
    """``BrOutsideARunTest.body(...)``: the tokens of a document holding `body_xml`, and its package."""
    pkg = package_of(body_xml, **options)
    out: list[str] = []
    walk(pkg.main_document_part.contents.body, out)
    return out, pkg


def package_of(body_xml: str, **options: Any) -> Any:
    pkg = create_package()
    if options:
        pkg.load_options = LoadOptions(**options)
    pkg.main_document_part.set_bytes(document_xml(body_xml))
    return pkg


def document_xml(body_xml: str) -> bytes:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
    ).encode()


def entries(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def saved(pkg: Any, **kwargs: Any) -> bytes:
    out = io.BytesIO()
    pkg.save(out, **kwargs)
    return out.getvalue()


def story_text(pkg: Any) -> str:
    return plain_text(extract(pkg, include="headers, footers"))


# ---------------------------------------------------------------------------
# the oracle's copy
# ---------------------------------------------------------------------------


def test_the_probes_are_docx4js():
    assert len(PROBES) == 13
    assert all(word(stem).exists() for stem in PROBES)
    listed = {}
    for line in (FIXTURES / "SHA1SUMS").read_text().splitlines():
        digest, name = line.split(maxsplit=1)
        listed[name] = digest
    on_disk = {p.name for p in FIXTURES.iterdir()} - {"SHA1SUMS", "README.md"}
    assert set(listed) == on_disk
    assert len(listed) == 27
    for name, digest in listed.items():
        assert hashlib.sha1((FIXTURES / name).read_bytes()).hexdigest() == digest, name


# ---------------------------------------------------------------------------
# each probe against Word's save
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("stem", SAME_AS_WORD)
def test_each_probe_reads_as_words_save(stem: str) -> None:
    assert tokens(load(source(stem))) == tokens(load(word(stem)))


@pytest.mark.parametrize("stem", SAME_AS_WORD)
def test_each_probe_has_words_plain_text(stem: str) -> None:
    assert story_text(load(source(stem))) == story_text(load(word(stem)))


def test_the_bare_instr_text_is_a_field_code_its_text_and_breaks_whole():
    """Q6 (docx4j's D6): Word saves the header's bare ``w:instrText`` as ``w:t``; here it stays a field code."""
    ours = tokens(load(source(BARE_INSTR_TEXT)))
    assert "instrText: I2 line one " in ours
    as_text = ["t" + t[len("instrText") :] if t.startswith("instrText") else t for t in ours]
    assert as_text == tokens(load(word(BARE_INSTR_TEXT)))


def test_the_bare_instr_texts_header_differs_in_plain_text_by_q6_alone():
    """The field code is hidden, so the header's plain text is its two breaks and its end; Word's shows the lines."""
    assert story_text(load(source(BARE_INSTR_TEXT))) == "\n\n\nI2 body text; the case is in the header\n"
    assert (
        story_text(load(word(BARE_INSTR_TEXT)))
        == " \n I2 line one \n I2 line two \nI2 body text; the case is in the header\n"
    )


def test_the_run_properties_are_words():
    """``br2-1``'s eight cases (docx4j's D1): R3 and R5 the mark's ``w:sz w:val="28"``, R6 its bold run's, the rest none."""
    ours = tokens(load(source("br2-1-run-properties")))
    assert "br[sz28]" in ours
    assert ours.count("br[sz28]") == 2
    assert ours.count("br[b sz40]") == 1
    assert ours.count("br") == 6


def test_a_field_code_keeps_its_break_and_updates_as_words_save():
    """``br2-2``: the code split round its break, whole, so ``fields: update`` evaluates all of it."""
    ours = load(source("br2-2-instrtext-field"))
    theirs = load(word("br2-2-instrtext-field"))
    assert plain_text(extract(ours, fields="update")) == "I1 start I1 code beforeI1 code after I1 end\n"
    assert plain_text(extract(ours, fields="update")) == plain_text(extract(theirs, fields="update"))


@pytest.mark.parametrize(
    ("stem", "text"),
    [
        ("br-2-p-between-runs", "BR2 before\x0bBR2 after"),
        ("br-3-t-mixed", "BR3 before\x0bBR3 after"),
        ("br-4-t-nested", "BR4 before\x0bBR4 after"),
    ],
)
def test_the_content_api_reads_the_break_and_all_the_text(stem: str, text: str) -> None:
    pkg = load(source(stem))
    assert text in [p.text for p in pkg.body.paragraphs]


# ---------------------------------------------------------------------------
# the records
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("stem", PROBES)
def test_each_part_records_its_warning_and_nothing_is_skipped(stem: str) -> None:
    pkg = load(source(stem))
    tokens(pkg)
    found = [(w.part_name, w.breaks, w.text_elements) for w in pkg.warnings]
    assert found == WARNINGS[stem]
    for part in story_parts(pkg):
        assert part.skipped == [], (part.part_name, part.skipped[:3])
        want = [(n, b, t) for n, b, t in WARNINGS[stem] if n == str(part.part_name)]
        assert [(w.part_name, w.breaks, w.text_elements) for w in part.warnings] == want
    assert pkg.skipped == []


def test_the_record_and_its_message():
    pkg = load(source("br2-1-run-properties"))
    _ = pkg.main_document_part.contents
    (record,) = pkg.main_document_part.warnings
    assert isinstance(record, BrOutsideRun)
    assert record.code == BrOutsideRun.code == "br.outside_run"
    assert record.message == (
        "/word/document.xml: w:br outside a w:r: 8 outside a run, 1 text elements holding elements; "
        "read as Word reads them"
    )
    assert "CR-073" not in record.message
    assert str(record) == record.message
    assert record.to_dict() == {
        "code": "br.outside_run",
        "part_name": "/word/document.xml",
        "breaks": 8,
        "text_elements": 1,
        "message": record.message,
    }
    with pytest.raises(AttributeError):
        record.breaks = 9  # type: ignore[misc]


def test_the_warning_is_logged_once_a_part(caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level(logging.WARNING, logger="docx4j_py.openpackaging"):
        pkg = load(source("br2-1-run-properties"))
        _ = pkg.main_document_part.contents
    ours = [r for r in caplog.records if r.name == "docx4j_py.openpackaging" and "w:br outside a w:r" in r.getMessage()]
    assert [r.getMessage() for r in ours] == [pkg.main_document_part.warnings[0].message]
    assert ours[0].levelno == logging.WARNING


@pytest.mark.parametrize("stem", PROBES)
def test_no_warning_on_words_saves(stem: str) -> None:
    pkg = load(word(stem))
    tokens(pkg)
    assert pkg.warnings == []
    assert pkg.skipped == []


def test_a_clean_part_has_no_warnings():
    pkg = load(SAMPLES / "2010-sample1.docx")
    pkg.unmarshal_all()
    assert pkg.warnings == []
    assert all(part.warnings == [] for part in pkg.parts.parts() if isinstance(part, XmlPart))


# ---------------------------------------------------------------------------
# strict, and the saves
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("stem", PROBES)
def test_strict_loads_every_probe(stem: str) -> None:
    """Q3: the pass runs before strict's test, and these are Word's reading of the file, not lenience."""
    pkg = load(source(stem), options=LoadOptions(strict=True))
    for part in story_parts(pkg):
        assert part.contents is not None
    assert [(w.part_name, w.breaks, w.text_elements) for w in pkg.warnings] == WARNINGS[stem]


@pytest.mark.parametrize("stem", PROBES)
def test_a_checked_save_reads_back_the_same_with_no_warning(stem: str) -> None:
    pkg = load(source(stem))
    before = tokens(pkg)
    again = load(saved(pkg, check=True))
    assert tokens(again) == before
    assert again.warnings == []
    assert again.skipped == []


@pytest.mark.parametrize("normalise", [True, False])
@pytest.mark.parametrize("stem", PROBES)
def test_a_saved_part_is_its_repaired_source(stem: str, normalise: bool) -> None:
    """Section 4.6's exception: identical to the source run through the same pass, the pass counted ``br-outside-run``."""
    pkg = load(source(stem))
    tokens(pkg)
    out = entries(saved(pkg, normalise=normalise))
    src = entries(source(stem).read_bytes())
    for part in story_parts(pkg):
        name = part.part_name.store_name
        report = canon.compare(src[name], out[name], normalise=normalise, repaired=True)
        assert report.identical, (name, report.differences, report.diffs[:5])
        counted = sum(w.breaks + w.text_elements for w in part.warnings)
        assert report.counts.get("br-outside-run", 0) == counted
        if counted:
            assert not canon.compare(src[name], out[name], normalise=normalise).identical


@pytest.mark.parametrize("stem", PROBES)
def test_an_untouched_save_is_byte_for_byte(stem: str) -> None:
    src = entries(source(stem).read_bytes())
    out = entries(saved(load(source(stem))))
    for name, data in src.items():
        if name == "[Content_Types].xml" or name.endswith(".rels"):
            continue
        assert out[name] == data, name


def test_repair_bytes_returns_the_bytes_themselves_where_nothing_changes():
    data = document_xml("<w:p><w:r><w:t>a</w:t><w:br/></w:r></w:p>")
    assert repair_bytes(data) == (data, 0, 0, 0)
    assert repair_bytes(data)[0] is data


# ---------------------------------------------------------------------------
# docx4j's edges (BrOutsideARunTest's five methods)
# ---------------------------------------------------------------------------


def test_legal_breaks_untouched():
    """A math run holds a ``w:br`` legally and keeps it while a stray one beside it is repaired; a ``w:r/w:br`` stays."""
    got, pkg = body_tokens(
        "<w:p><w:r><w:t>a</w:t><w:br/></w:r>"
        "<m:oMath><m:r><w:br/><m:t>x</m:t></m:r></m:oMath>"
        "<w:br/><w:r><w:t>b</w:t></w:r></w:p>"
    )
    assert got == ["t:a", "br", "br", "t:b", "¶"]
    paragraph = pkg.main_document_part.contents.body.content[0]
    maths = [child for _q, child in iter_children(paragraph) if type(child).__name__ == "OMath"]
    assert len(maths) == 1
    (math_run,) = [child for _q, child in iter_children(maths[0])]
    assert any(type(child).__name__ == "Br" for _q, child in iter_children(math_run))
    assert [(w.breaks, w.text_elements) for w in pkg.warnings] == [(1, 0)]


def test_whitespace_in_text_without_preserve():
    assert body_tokens("<w:p><w:r><w:t>  <w:br/></w:t></w:r></w:p>")[0] == ["br", "¶"]
    assert body_tokens('<w:p><w:r><w:t xml:space="preserve">  <w:br/></w:t></w:r></w:p>')[0] == ["t:  ", "br", "¶"]


def test_mark_properties_without_the_marks_own():
    _got, pkg = body_tokens(
        '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="a"/><w:b/></w:rPr></w:pPr><w:br/><w:r><w:t>x</w:t></w:r></w:p>'
    )
    run = pkg.main_document_part.contents.body.content[0].content[0]
    assert type(run).__name__ == "R"
    assert type(run.content[0]).__name__ == "Br"
    assert run.r_pr.b is not None
    assert "ins" not in {name.rpartition("}")[2] for name, _c in iter_children(run.r_pr)}


def test_deleted_text():
    assert body_tokens('<w:p><w:del w:id="1" w:author="a"><w:r><w:delText>x<w:br/>y</w:delText></w:r></w:del></w:p>')[
        0
    ] == ["delText:x", "br", "delText:y", "¶"]
    assert body_tokens(
        '<w:p><w:del w:id="1" w:author="a"><w:r><w:delText>x</w:delText></w:r><w:br/>'
        "<w:r><w:delText>y</w:delText></w:r></w:del></w:p>"
    )[0] == ["delText:x", "br", "delText:y", "¶"]


def test_body_break_without_paragraph_after():
    got, pkg = body_tokens(
        "<w:p><w:r><w:t>a</w:t></w:r></w:p><w:br/><w:br/><w:p><w:r><w:t>b</w:t></w:r></w:p>"
        '<w:br w:type="page"/><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
        "<w:tr><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    )
    assert got == ["t:a", "¶", "br", "br", "t:b", "¶", "br(page)", "¶", "t:c", "¶"]
    content = pkg.main_document_part.contents.body.content
    assert [type(c).__name__ for c in content] == ["P", "P", "P", "Tbl"]
    assert [(w.breaks, w.text_elements) for w in pkg.warnings] == [(3, 0)]


# ---------------------------------------------------------------------------
# this package's own edges (section 2.6)
# ---------------------------------------------------------------------------


def test_a_br_in_no_namespace_keeps_the_text_whole_with_no_break():
    """docx4j's survey sample ``01af91808eac``: an element in a ``w:t`` gives its text, and a ``br`` is not ``w:br``."""
    got, pkg = body_tokens("<w:p><w:r><w:t>a &lt;br /&gt;<br/>b</w:t></w:r></w:p>")
    assert got == ["t:a <br />b", "¶"]
    assert [(w.breaks, w.text_elements) for w in pkg.warnings] == [(0, 1)]
    assert pkg.main_document_part.skipped == []


def test_a_stretch_of_only_no_break_spaces_is_text():
    """``normalize-space`` treats only XML's four whitespace characters as space, unlike ``str.strip``."""
    assert body_tokens("<w:p><w:r><w:t>\u00a0<w:br/>b</w:t></w:r></w:p>")[0] == ["t:\u00a0", "br", "t:b", "¶"]


def test_a_nested_text_element_holding_a_break_counts_twice():
    got, pkg = body_tokens("<w:p><w:r><w:t><w:t>a<w:br/>b</w:t></w:t></w:r></w:p>")
    assert got == ["t:a", "br", "t:b", "¶"]
    assert [(w.breaks, w.text_elements) for w in pkg.warnings] == [(0, 2)]


def test_a_piece_takes_the_nearest_xml_space():
    """Lines 886 to 897: the nearest ancestor's ``xml:space``, else none."""
    data, _elements, _breaks, texts = repair_bytes(
        document_xml('<w:p><w:r><w:t xml:space="preserve"><w:t> a </w:t><w:br/></w:t><w:t>b<w:br/>c</w:t></w:r></w:p>')
    )
    assert texts == 2
    assert b'<w:t xml:space="preserve"> a </w:t><w:br/><w:t>b</w:t><w:br/><w:t>c</w:t>' in data
    assert body_tokens('<w:p><w:r><w:t xml:space="preserve"><w:t> a </w:t><w:br/></w:t></w:r></w:p>')[0] == [
        "t: a ",
        "br",
        "¶",
    ]


def test_a_tab_in_a_text_element_keeps_the_text_either_side_and_loses_the_tab():
    """Word draws no tab either (check 76's ``76z``), so the tab is gone and the part says so (section 11)."""
    got, pkg = body_tokens("<w:p><w:r><w:t>a<w:tab/>b</w:t></w:r></w:p>")
    assert got == ["t:ab", "¶"]
    assert [(w.breaks, w.text_elements) for w in pkg.warnings] == [(0, 1)]
    assert [item.qname for item in pkg.main_document_part.skipped] == [W + "tab"]


def test_a_break_in_a_block_level_sdt_content_is_counted_left_and_skipped():
    got, pkg = body_tokens("<w:sdt><w:sdtContent><w:br/><w:p><w:r><w:t>a</w:t></w:r></w:p></w:sdtContent></w:sdt>")
    assert got == ["t:a", "¶"]
    part = pkg.main_document_part
    assert [(w.breaks, w.text_elements) for w in part.warnings] == [(1, 0)]
    assert len(part.skipped) == 1 and part.skipped[0].qname == W + "br"
    strict = package_of(
        "<w:sdt><w:sdtContent><w:br/><w:p><w:r><w:t>a</w:t></w:r></w:p></w:sdtContent></w:sdt>", strict=True
    )
    with pytest.raises(Docx4JException, match="strict=True"):
        _ = strict.main_document_part.contents


KEPT_BRANCH = (
    '<mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>a</w:t></w:r><w:br/><w:r><w:t>b</w:t></w:r></w:p>'
    "</mc:Choice><mc:Fallback><w:p><w:r><w:t>f</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>"
)


def test_a_kept_branch_is_repaired_in_place():
    """``preserve``: a typed paragraph in a kept block-level ``mc:Choice`` is repaired there, and saved there."""
    got, pkg = body_tokens(KEPT_BRANCH)
    assert got == ["t:a", "br", "t:b", "¶"]
    xml = pkg.main_document_part.xml.decode()
    choice = xml[xml.index("<mc:Choice") : xml.index("</mc:Choice>")]
    assert "<w:r><w:br/></w:r>" in choice
    assert "<mc:Fallback>" in xml


@pytest.mark.parametrize("options", [{"mc_load_mode": "word"}, {"mce_preprocess": True}])
def test_a_branch_reads_the_same_in_each_mce_mode(options: dict[str, Any]) -> None:
    got, pkg = body_tokens(KEPT_BRANCH, **options)
    assert got == ["t:a", "br", "t:b", "¶"]
    assert [(w.breaks, w.text_elements) for w in pkg.warnings] == [(1, 0)]
    assert pkg.main_document_part.skipped == []


RUN_LEVEL_BRANCH = (
    '<w:p><w:r><w:t>a</w:t></w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:br/></mc:Choice>'
    "<mc:Fallback/></mc:AlternateContent><w:r><w:t>b</w:t></w:r></w:p>"
)


def test_a_break_misplaced_only_once_its_branch_is_selected_is_repaired_under_mce_preprocess_alone():
    """Section 4.7's departure: the stylesheet matches its input, so it leaves ``w:p/mc:AlternateContent/mc:Choice/w:br``.

    In ``preserve`` mode the branch's ``w:br`` is typed there (CR-012 Phase B), not skipped, so nothing triggers; under
    ``mce_preprocess=True`` the pass sees the resolved bytes, where it is directly in the paragraph, and repairs it.
    """
    _got, kept = body_tokens(RUN_LEVEL_BRANCH)
    assert kept.warnings == [] and kept.main_document_part.skipped == []
    assert kept.body.paragraphs[0].text == "ab"
    _got, resolved = body_tokens(RUN_LEVEL_BRANCH, mce_preprocess=True)
    assert [(w.breaks, w.text_elements) for w in resolved.warnings] == [(1, 0)]
    assert resolved.body.paragraphs[0].text == "a\x0bb"


def test_two_breaks_first_in_a_paragraph_both_take_the_marks_properties():
    """Decided on the tree as read: a run rule 1 makes is never a run before another break."""
    got, _pkg = body_tokens(
        '<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:br/><w:br/><w:r><w:t>x</w:t></w:r><w:br/></w:p>'
    )
    assert got == ["br[sz28]", "br[sz28]", "t:x", "br", "¶"]


def test_body_breaks_take_the_next_paragraphs_mark_after_its_ppr():
    got, pkg = body_tokens(
        '<w:br/><w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
    )
    assert got == ["br[b]", "t:x", "¶"]
    paragraph = pkg.main_document_part.contents.body.content[0]
    assert paragraph.p_pr is not None
    assert type(paragraph.content[0]).__name__ == "R"


# ---------------------------------------------------------------------------
# the routes
# ---------------------------------------------------------------------------

#: A cell whose ``w:t`` holds a break and whose last child is a table, so it does not end with a paragraph.
CELL = (
    '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>'
    "<w:p><w:r><w:t>a<w:br/>b</w:t></w:r></w:p>"
    '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
    "</w:tc></w:tr></w:tbl><w:p/>"
)


def cell_document() -> bytes:
    base = entries(saved(create_package()))
    base["word/document.xml"] = document_xml(CELL)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as archive:
        for name, data in base.items():
            archive.writestr(name, data)
    return out.getvalue()


def test_source_tree_reads_as_the_part_does_so_a_carried_finding_stays_carried():
    """Section 4.8: the save's check compares with ``source_tree()``, which goes through the same pass."""
    pkg = load(cell_document())
    _ = pkg.main_document_part.contents
    pkg.save(io.BytesIO(), check=True)
    assert "cell.no_end_paragraph" in [f.code for f in pkg.last_save.findings]
    assert pkg.main_document_part.source_tree() is not None
    assert pkg.main_document_part.warnings[0].text_elements == 1


def test_without_the_pass_in_source_tree_the_carried_finding_would_look_introduced(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Why section 4.8 holds: the cell's text differs (``a\\nb`` against ``a``), so the save would raise."""
    from docx4j_py.model.content import SaveCheckError
    from docx4j_py.runtime import parser

    def unrepaired(self: XmlPart) -> Any:
        return parser(lenient=True, skipped_report=True).from_bytes(self._source_bytes(), self.model_class)

    pkg = load(cell_document())
    _ = pkg.main_document_part.contents
    monkeypatch.setattr(XmlPart, "source_tree", unrepaired)
    with pytest.raises(SaveCheckError):
        pkg.save(io.BytesIO(), check=True)


def test_a_private_read_then_contents_adopts_the_repaired_tree_with_its_warning(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = []
    original = XmlPart._unmarshal

    def counting(part: XmlPart, private: bool = False) -> None:
        if str(part.part_name) == "/word/document.xml":
            calls.append(private)
        original(part, private)

    monkeypatch.setattr(XmlPart, "_unmarshal", counting)
    part = load(source("br-2-p-between-runs")).main_document_part
    private = part.read_contents()
    assert not part.is_unmarshalled
    assert [(w.breaks, w.text_elements) for w in part.warnings] == [(1, 0)]
    assert part.contents is private
    assert calls == [True], "adopted, not parsed again"
    assert [(w.breaks, w.text_elements) for w in part.warnings] == [(1, 0)]
    assert part.skipped == []


def test_set_bytes_clears_the_warning():
    pkg = load(source("br-1-p-alone"))
    part = pkg.main_document_part
    _ = part.contents
    assert part.warnings
    part.set_bytes(document_xml("<w:p><w:r><w:t>clean</w:t></w:r></w:p>"))
    assert part.warnings == []
    _ = part.contents
    assert part.warnings == [] and pkg.warnings == []


def test_insert_ooxml_reads_a_fragment_as_word_reads_it_and_reports():
    """Q4: ``insert_ooxml``'s fragment route repairs, and the call's report says so."""
    pkg = create_package()
    views = pkg.body.insert_ooxml(
        "<w:p><w:r><w:t>a</w:t></w:r><w:br/><w:r><w:t>b<w:br/>c</w:t></w:r></w:p>"
        "<w:br/><w:p><w:r><w:t>d</w:t></w:r></w:p>"
    )
    assert [view.text for view in views] == ["a\x0bb\x0bc", "\x0bd"]
    report = pkg.last_change
    assert report.warnings == (
        "the fragment: w:br outside a w:r: 2 outside a run, 1 text elements holding elements; read as Word reads them",
    )
    again = load(saved(pkg, check=True))
    assert [p.text for p in again.body.paragraphs][-2:] == ["a\x0bb\x0bc", "\x0bd"]
    assert again.warnings == []


def test_insert_ooxml_of_a_clean_fragment_warns_nothing():
    pkg = create_package()
    pkg.body.insert_ooxml("<w:p><w:r><w:t>a</w:t><w:br/><w:t>b</w:t></w:r></w:p>")
    assert pkg.last_change.warnings == ()


def test_wml_still_refuses_a_break_outside_a_run():
    """Q4: ``wml()`` is code's own builder, and its error names the skipped ``w:br``."""
    with pytest.raises(FragmentError, match="br"):
        wml("<w:p><w:br/></w:p>")
    with pytest.raises(FragmentError):
        wml("<w:p><w:r><w:t>a<w:br/>b</w:t></w:r></w:p>")
    with pytest.raises(FragmentError):
        wml.all("<w:p><w:br/></w:p>", wrapper="body")


# ---------------------------------------------------------------------------
# the trigger, and its cost
# ---------------------------------------------------------------------------


def _first_parse_report(data: bytes, clazz: type) -> list[Any]:
    from docx4j_py.runtime import parser

    engine = parser(lenient=True, skipped_report=True)
    engine.from_bytes(data, clazz)
    return list(engine.skipped.items)


@pytest.mark.parametrize("stem", PROBES)
def test_every_probe_triggers_by_the_reports_path(stem: str) -> None:
    """Risk 1: a change to the report's path would silence the pass; the thirteen pin it."""
    pkg = load(source(stem))
    affected = [part for part in story_parts(pkg) if any(n == str(part.part_name) for n, _b, _t in WARNINGS[stem])]
    assert affected
    for part in affected:
        report = _first_parse_report(part._raw_bytes(), part.model_class)
        assert report
        assert triggers(report)


def test_parent_qname_is_a_class_name_for_a_choice_class_so_the_path_is_read():
    """Section 2.3: the report names a ``w:t``'s class (``{w}RT``) as the parent; the path names the element."""
    pkg = load(source("br-3-t-mixed"))
    part = pkg.main_document_part
    (item,) = _first_parse_report(part._raw_bytes(), part.model_class)
    assert item.parent_qname == W + "RT"
    assert item.path.endswith(f"{W}r/{W}t/{W}br")
    assert triggers([item])


def test_an_attribute_or_a_legal_parent_never_triggers():
    from types import SimpleNamespace

    attribute = SimpleNamespace(kind="attribute", qname=W + "foo", path=f"{W}document/{W}body/{W}p/{W}r/{W}t/@{W}foo")
    in_run = SimpleNamespace(kind="element", qname=W + "br", path=f"{W}document/{W}body/{W}p/{W}r/{W}br")
    in_math = SimpleNamespace(kind="element", qname=W + "foo", path=f"{W}document/{W}body/{W}p/{{http://x}}r/{W}foo")
    assert not triggers([attribute, in_run, in_math])
    assert not triggers([])


def test_a_clean_part_is_parsed_once_and_the_pass_never_runs(monkeypatch: pytest.MonkeyPatch) -> None:
    """Section 5's cost, held deterministically: Symbols' main part (790,857 bytes, two legal breaks)."""
    from docx4j_py import runtime

    parses = []
    real_parser = runtime.parser

    def counting_parser(**kwargs: Any) -> Any:
        engine = real_parser(**kwargs)
        real_from_bytes = engine.from_bytes

        def from_bytes(*args: Any, **kw: Any) -> Any:
            parses.append(1)
            return real_from_bytes(*args, **kw)

        engine.from_bytes = from_bytes  # type: ignore[method-assign]
        return engine

    passes = []

    def never(*args: Any, **kwargs: Any) -> Any:
        passes.append(args)
        raise AssertionError("the pass ran on a clean part")

    pkg = load(SAMPLES / "Symbols.docx")
    part = pkg.main_document_part
    monkeypatch.setattr(runtime, "parser", counting_parser)
    monkeypatch.setattr(breaks, "triggers", never)
    monkeypatch.setattr(read_through, "repair_bytes", never)
    assert part.contents is not None
    assert len(part._source_bytes()) == 790_857
    assert parses == [1]
    assert passes == []
    assert part.warnings == [] and part.skipped == []


def test_an_affected_part_is_parsed_twice(monkeypatch: pytest.MonkeyPatch) -> None:
    from docx4j_py import runtime

    parses = []
    real_parser = runtime.parser

    def counting_parser(**kwargs: Any) -> Any:
        parses.append(1)
        return real_parser(**kwargs)

    part = load(source("br2-1-run-properties")).main_document_part
    monkeypatch.setattr(runtime, "parser", counting_parser)
    _ = part.contents
    assert parses == [1, 1]


# ---------------------------------------------------------------------------
# an element Word doesn't know in a text element (section 11, core-ts's check 76)
# ---------------------------------------------------------------------------

CHECK76 = ROOT / "tests" / "fixtures" / "check76"
EXT = 'xmlns:ext="urn:ext"'
MATH = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
OPAQUE = "{urn:docx4j:check76}opaque"
#: Each file: what Word 365 reads in its middle paragraph, and the element that goes.
CHECK76_FILES = {
    "76m-ignorable-in-t": ("TARGET TEXT", OPAQUE),
    "76n-ignorable-with-text-in-t": ("TARGET TEXT", OPAQUE),
    "76u-madeup-with-text-in-t": ("TARGET TEXT", W + "madeUp"),
    "76w-ignorable-holding-a-break-in-t": ("TARGET TEXT", OPAQUE),
    "76z-tab-in-t": ("TARGETTEXT", W + "tab"),
}


def test_the_check_76_files_are_core_tss():
    listed = dict(reversed(line.split()) for line in (CHECK76 / "SHA1SUMS").read_text().splitlines())
    assert set(listed) == {path.name for path in CHECK76.glob("*.docx")}
    assert len(listed) == 144
    for name, digest in listed.items():
        assert hashlib.sha1((CHECK76 / name).read_bytes()).hexdigest() == digest, name


@pytest.mark.parametrize("stem", CHECK76_FILES)
def test_each_check_76_file_reads_as_words_save(stem: str) -> None:
    ours, words = load(CHECK76 / f"{stem}.docx"), load(CHECK76 / f"{stem}-word365.docx")
    assert ours.body.paragraphs[1].text == CHECK76_FILES[stem][0]
    assert tokens(ours) == tokens(words)
    assert story_text(ours) == story_text(words)
    assert words.warnings == [] and words.main_document_part.skipped == []


@pytest.mark.parametrize("stem", CHECK76_FILES)
def test_each_check_76_file_says_what_went(stem: str) -> None:
    pkg = load(CHECK76 / f"{stem}.docx")
    part = pkg.main_document_part
    assert part.contents is not None
    assert [(w.breaks, w.text_elements) for w in part.warnings] == [(0, 1)]
    assert [(item.kind, item.qname, item.parent_qname) for item in part.skipped] == [
        ("element", CHECK76_FILES[stem][1], W + "t")
    ]
    assert [item for _name, item in pkg.skipped] == part.skipped


@pytest.mark.parametrize("stem", CHECK76_FILES)
def test_a_save_of_each_check_76_file_writes_what_word_reads(stem: str) -> None:
    pkg = load(CHECK76 / f"{stem}.docx")
    assert pkg.main_document_part.contents is not None
    data = saved(pkg)
    xml = entries(data)["word/document.xml"]
    assert b"<ext:opaque" not in xml and b"<w:madeUp" not in xml and b"<w:tab" not in xml
    assert b">HIDDEN WORDS<" not in xml
    again = load(data)
    assert tokens(again) == tokens(load(CHECK76 / f"{stem}-word365.docx"))
    assert again.warnings == [] and again.main_document_part.skipped == []


@pytest.mark.parametrize("normalise", [True, False])
def test_a_saved_check_76_part_is_its_repaired_source(normalise: bool) -> None:
    source_path = CHECK76 / "76n-ignorable-with-text-in-t.docx"
    pkg = load(source_path)
    assert pkg.main_document_part.contents is not None
    data = entries(saved(pkg, normalise=normalise))["word/document.xml"]
    report = canon.compare(entries(source_path.read_bytes())["word/document.xml"], data, repaired=True, normalise=normalise)
    assert report.identical, (report.differences, report.diffs[:5])


@pytest.mark.parametrize(
    ("text_element", "tokens_wanted"),
    [
        (f"<w:t>a<ext:o {EXT}>b<w:br/><w:t>c</w:t><ext:deeper>d</ext:deeper></ext:o>e</w:t>", ["t:ae", "¶"]),
        (f"<w:t>a<ext:o {EXT}>b</ext:o><w:br/>c</w:t>", ["t:a", "br", "t:c", "¶"]),
        ("<w:t>a<w:madeUp>b<w:br/><w:t>c</w:t></w:madeUp>d</w:t>", ["t:ad", "¶"]),
        (f"<w:t><ext:o {EXT}>only</ext:o></w:t>", ["¶"]),
        (f"<w:t>a<w:t>b<ext:o {EXT}>c</ext:o>d</w:t>e</w:t>", ["t:abde", "¶"]),
        ("<w:t>a<w:r><w:t>b</w:t></w:r>c</w:t>", ["t:ac", "¶"]),
    ],
)
def test_all_an_element_holds_goes_with_it(text_element: str, tokens_wanted: list[str]) -> None:
    assert body_tokens(f"<w:p><w:r>{text_element}</w:r></w:p>")[0] == tokens_wanted


def test_deleted_text_and_a_field_code_lose_it_too():
    deleted = f'<w:p><w:del w:id="1" w:author="a"><w:r><w:delText>a<ext:o {EXT}>b</ext:o>c</w:delText></w:r></w:del></w:p>'
    assert body_tokens(deleted)[0] == ["delText:ac", "¶"]
    _got, pkg = body_tokens(f"<w:p><w:r><w:instrText>a<ext:o {EXT}>b</ext:o>c</w:instrText></w:r></w:p>")
    assert b">b<" not in entries(saved(pkg))["word/document.xml"]


@pytest.mark.parametrize(
    ("text_element", "tokens_wanted"),
    [
        ("<w:t>a<span>b</span>c</w:t>", ["t:abc", "¶"]),
        ("<w:t>a<span>b<w:br/>c</span>d</w:t>", ["t:ab", "br", "t:cd", "¶"]),
        (f"<w:t>a<m:t {MATH}>b</m:t>c</w:t>", ["t:abc", "¶"]),
    ],
)
def test_no_namespace_and_a_namespace_word_knows_give_their_text_as_docx4j(
    text_element: str, tokens_wanted: list[str]
) -> None:
    """Word refuses the first two (``76x``, ``76y``), and the third is unmeasured: nothing is lost, so nothing is skipped."""
    got, pkg = body_tokens(f"<w:p><w:r>{text_element}</w:r></w:p>")
    assert got == tokens_wanted
    assert pkg.main_document_part.skipped == []


def test_each_element_taken_is_named_where_the_forks_report_stops_short():
    """The fork names the outer ``w:t`` of two; the pass names what went, in the report's own form."""
    _got, pkg = body_tokens(f"<w:p><w:r><w:t>a<w:t>b<ext:o {EXT}>c</ext:o>d</w:t>e<w:tab/></w:t></w:r></w:p>")
    first, second = pkg.main_document_part.skipped
    assert (first.qname, first.parent_qname) == ("{urn:ext}o", W + "t")
    assert first.path == "/".join(W + step for step in ("document", "body", "p", "r", "t", "t")) + "/{urn:ext}o"
    assert second.path == "/".join(W + step for step in ("document", "body", "p", "r", "t", "tab"))


def test_what_the_second_parse_skips_stays_beside_what_the_pass_took():
    _got, pkg = body_tokens("<w:p><w:madeUp/><w:r><w:t>a<w:madeUp/>b</w:t></w:r></w:p>")
    assert [item.parent_qname for item in pkg.main_document_part.skipped] == [W + "p", W + "t"]


@pytest.mark.parametrize("stem", CHECK76_FILES)
def test_strict_refuses_a_part_the_pass_took_something_from(stem: str) -> None:
    """Q3's strict load repairs, but no more than that: what went is skipped content."""
    with pytest.raises(Docx4JException, match="strict=True"):
        assert load(CHECK76 / f"{stem}.docx", options=LoadOptions(strict=True)).main_document_part.contents is not None


def test_insert_ooxml_refuses_a_fragment_it_would_take_something_from():
    """Any loss in a fragment is refused, a text element's as the parse's own."""
    pkg = create_package()
    before = len(pkg.body.paragraphs)
    with pytest.raises(FragmentError, match="1 item"):
        pkg.body.insert_ooxml(f"<w:p><w:r><w:t>a<ext:o {EXT}>b</ext:o>c</w:t></w:r></w:p>")
    assert len(pkg.body.paragraphs) == before


def test_a_lenient_fragment_accepts_the_loss():
    found: list[Any] = []
    (paragraph,) = wml.all(
        f"<w:p><w:r><w:t>a<ext:o {EXT}>b</ext:o>c</w:t></w:r></w:p>", wrapper="body", lenient=True, _breaks=found
    )
    assert [(record.breaks, record.text_elements) for record in found] == [(0, 1)]
    assert text_of(paragraph) == "ac"
