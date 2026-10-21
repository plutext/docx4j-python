"""``mc:AlternateContent`` at block, row and cell level: CR-012 Phase A, docx4j CR-051 phase 1.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_mce_levels.py -q

The probes are docx4j's ``mc-fun`` fixtures at ``2853b7280`` (Jason's, with
Word 365's saves of them), copied byte for byte into ``tests/fixtures/mc-fun/``
(``tests/README.md`` has the provenance). Most carry ten Choices, newest first
(``w16sdtfl`` to ``w14``), and a Fallback; Word 365 reads the first it
understands. The built cases are docx4j's ``McBlockLevelLoadTest``'s: an element
of three branches (a Choice requiring ``wpi``, which docx4j's list does not
understand, one requiring ``w14``, which it does, and a Fallback) in each
container the schema admits it in since the refresh.

What is held (CR-012 section 5, Phase A):

* no probe skips anything, in either mode, and ``strict=True`` loads each;
* ``preserve`` (the default) saves each probe's main part canonically identical
  to its source, every branch kept, nothing recorded;
* ``word`` keeps the branch Word 365 keeps, and a save describes as Word's save
  of the probe does (``tests/revisions/description.py``), text boxes apart;
* the splice in every container, REQ-305, an element with no branch, nesting,
  ``mc:Ignorable`` and the record, in ``word`` mode;
* a ``w:sdt`` straight in a branch typed as its container's kind (the fork's
  stage 11), round-tripping whole;
* ``mce_preprocess`` over docx4j's list (Q8), and the mode as a load option and
  a package setting;
* the readers' audit (section 4.1 item 5), one test per reader, in ``preserve``
  mode; the revisions twins are ``tests/revisions/test_alternate_content.py``.
"""

from __future__ import annotations

import io
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from canon import normalize

from docx4j_py import load
from docx4j_py.fragments import FragmentError, to_xml, wml
from docx4j_py.mce import AlternateContent, AlternateContentChoice
from docx4j_py.model.content.errors import AddressError
from docx4j_py.model.text import extract, plain_text
from docx4j_py.openpackaging import LoadOptions
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.mc_block_level import (
    ResolvedAlternateContent,
    splice_word,
)
from docx4j_py.openpackaging.parts.wml import (
    CommentsPart,
    EndnotesPart,
    FooterPart,
    FootnotesPart,
    GlossaryDocumentPart,
    HeaderPart,
)
from docx4j_py.traversal import find, text_of
from docx4j_py.wml import (
    CTCustomXmlBlock,
    CTSdtCell,
    CTSdtRow,
    CTTxbxContent,
    P,
    SdtBlock,
    SdtRun,
    Tc,
    Tr,
)
from tests.revisions.description import describe

PROBES_DIR = ROOT / "tests" / "fixtures" / "mc-fun"
PROBES = sorted(
    p.stem for p in PROBES_DIR.glob("*.docx") if not p.stem.endswith("-word365")
)
WITH_WORD_SAVE = sorted(
    stem for stem in PROBES if (PROBES_DIR / f"{stem}-word365.docx").exists()
)
MODES = ("preserve", "word")

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
NS = (
    f'xmlns:w="{_W}"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    f' xmlns:mc="{_MC}"'
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
    ' xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"'
    ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
)


def probe(stem: str, mode: str = "preserve") -> Any:
    """A probe, loaded in a mode, its main part unmarshalled."""
    pkg = load(PROBES_DIR / f"{stem}.docx", options=LoadOptions(mc_block_level=mode))
    assert pkg.main_document_part.contents is not None
    return pkg


def wml_parts(pkg: Any) -> list[Any]:
    return [
        part
        for part in pkg.parts.values()
        if getattr(part, "model_class_path", None)
        and part.model_class_path.startswith("docx4j_py.wml.")
    ]


def text(pkg: Any) -> str:
    return plain_text(extract(pkg))


def main_xml(data: bytes) -> bytes:
    return zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml")


def count(xml: bytes, token: str) -> int:
    return xml.decode("utf-8").count(token)


# ---------------------------------------------------------------------------
# a package from raw XML, with the parts a container needs (docx4j McProbeDocx)
# ---------------------------------------------------------------------------

_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml."
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


class Probe:
    """A ``.docx`` made from raw XML: a body, and parts related from the main document part."""

    def __init__(self, body: str, root_attributes: str = "") -> None:
        self.parts = {
            "word/document.xml": f"{_DECLARATION}<w:document {NS}{root_attributes}><w:body>{body}</w:body></w:document>"
        }
        self.types: list[str] = []
        self.rels: list[str] = []

    def part(
        self, name: str, kind: str, rel: str, rel_id: str, root: str, inner: str
    ) -> Probe:
        self.parts[f"word/{name}"] = f"{_DECLARATION}<{root} {NS}>{inner}</{root}>"
        self.types.append(
            f'<Override PartName="/word/{name}" ContentType="{_CT}{kind}+xml"/>'
        )
        self.rels.append(
            f'<Relationship Id="{rel_id}" Type="{_REL}{rel}" Target="{name}"/>'
        )
        return self

    def save(self) -> bytes:
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr(
                "[Content_Types].xml",
                f'{_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                f'<Override PartName="/word/document.xml" ContentType="{_CT}document.main+xml"/>'
                f"{''.join(self.types)}</Types>",
            )
            z.writestr(
                "_rels/.rels",
                f'{_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{_REL}officeDocument" Target="word/document.xml"/></Relationships>',
            )
            z.writestr(
                "word/_rels/document.xml.rels",
                f'{_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f"{''.join(self.rels)}</Relationships>",
            )
            for name, data in self.parts.items():
                z.writestr(name, data)
        return out.getvalue()

    def load(self, mode: str = "preserve", **options: Any) -> Any:
        return load(self.save(), options=LoadOptions(mc_block_level=mode, **options))


def p(text_: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{text_}</w:t></w:r></w:p>'


def tc(text_: str) -> str:
    return f"<w:tc>{p(text_)}</w:tc>"


def tr(text_: str) -> str:
    return f"<w:tr>{tc(text_)}</w:tr>"


TBL_PROPS = '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>'


def tbl(rows: str) -> str:
    return f"<w:tbl>{TBL_PROPS}{rows}</w:tbl>"


def ac(kind: str, tag: str) -> str:
    """docx4j's three-branch element: a ``wpi`` Choice (not understood), a ``w14`` Choice (understood), a Fallback."""
    block = {"row": tr, "cell": tc}.get(kind, p)
    return (
        "<mc:AlternateContent>"
        f'<mc:Choice Requires="wpi">{block(tag + "_WPI")}</mc:Choice>'
        f'<mc:Choice Requires="w14">{block(tag + "_W14")}</mc:Choice>'
        f"<mc:Fallback>{block(tag + '_FB')}</mc:Fallback>"
        "</mc:AlternateContent>"
    )


def text_box(blocks: str) -> str:
    """A bare wps text box in a run (Word 2010 and later read it without alternate content)."""
    return (
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><wp:docPr id="1" name="Text Box 1"/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>'
        f"<wps:txbx><w:txbxContent>{blocks}</w:txbxContent></wps:txbx>"
        "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )


SECT_PR = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/></w:sectPr>'


def _first(root: Any, cls: type) -> Any:
    (found, *_) = find(root, cls, mce="all")
    return found


def _part(pkg: Any, cls: type) -> Any:
    return next(part for part in pkg.parts.values() if isinstance(part, cls))


def _txbx(pkg: Any) -> list:
    return _first(pkg.main_document_part.contents, CTTxbxContent).content


#: docx4j McBlockLevelLoadTest.cases(): each container, the element at index 1 of its content.
CASES: dict[str, tuple[str, Any, Any, Any]] = {
    # name: (kind, the probe, the part, the content list)
    "w:body": (
        "block",
        lambda: Probe(p("before") + ac("block", "BODY") + p("after")),
        lambda pkg: pkg.main_document_part,
        lambda pkg: pkg.main_document_part.contents.body.content,
    ),
    "w:hdr": (
        "block",
        lambda: (
            Probe(p("body") + SECT_PR)
            .part(
                "header1.xml",
                "header",
                "header",
                "rIdH",
                "w:hdr",
                p("before") + ac("block", "HDR") + p("after"),
            )
            .part("footer1.xml", "footer", "footer", "rIdF", "w:ftr", p("footer"))
        ),
        lambda pkg: _part(pkg, HeaderPart),
        lambda pkg: _part(pkg, HeaderPart).contents.content,
    ),
    "w:ftr": (
        "block",
        lambda: (
            Probe(p("body") + SECT_PR)
            .part("header1.xml", "header", "header", "rIdH", "w:hdr", p("header"))
            .part(
                "footer1.xml",
                "footer",
                "footer",
                "rIdF",
                "w:ftr",
                p("before") + ac("block", "FTR") + p("after"),
            )
        ),
        lambda pkg: _part(pkg, FooterPart),
        lambda pkg: _part(pkg, FooterPart).contents.content,
    ),
    "w:tc": (
        "block",
        lambda: Probe(
            tbl(
                f"<w:tr><w:tc>{p('before')}{ac('block', 'TC')}{p('after')}</w:tc></w:tr>"
            )
        ),
        lambda pkg: pkg.main_document_part,
        lambda pkg: _first(pkg.main_document_part.contents, Tc).content,
    ),
    "w:footnote": (
        "block",
        lambda: Probe('<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>').part(
            "footnotes.xml",
            "footnotes",
            "footnotes",
            "rIdFn",
            "w:footnotes",
            f'<w:footnote w:id="1">{p("before")}{ac("block", "FN")}{p("after")}</w:footnote>',
        ),
        lambda pkg: _part(pkg, FootnotesPart),
        lambda pkg: _part(pkg, FootnotesPart).contents.footnote[0].content,
    ),
    "w:endnote": (
        "block",
        lambda: Probe('<w:p><w:r><w:endnoteReference w:id="1"/></w:r></w:p>').part(
            "endnotes.xml",
            "endnotes",
            "endnotes",
            "rIdEn",
            "w:endnotes",
            f'<w:endnote w:id="1">{p("before")}{ac("block", "EN")}{p("after")}</w:endnote>',
        ),
        lambda pkg: _part(pkg, EndnotesPart),
        lambda pkg: _part(pkg, EndnotesPart).contents.endnote[0].content,
    ),
    "w:comment": (
        "block",
        lambda: Probe(
            '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="1"/>'
            '<w:r><w:commentReference w:id="1"/></w:r></w:p>'
        ).part(
            "comments.xml",
            "comments",
            "comments",
            "rIdC",
            "w:comments",
            f'<w:comment w:id="1" w:author="a">{p("before")}{ac("block", "CMT")}{p("after")}</w:comment>',
        ),
        lambda pkg: _part(pkg, CommentsPart),
        lambda pkg: _part(pkg, CommentsPart).contents.comment[0].content,
    ),
    "w:txbxContent": (
        "block",
        lambda: Probe(text_box(p("before") + ac("block", "TXBX") + p("after"))),
        lambda pkg: pkg.main_document_part,
        _txbx,
    ),
    "w:sdtContent (block)": (
        "block",
        lambda: Probe(
            f"<w:sdt><w:sdtPr/><w:sdtContent>{p('before')}{ac('block', 'SDTB')}{p('after')}</w:sdtContent></w:sdt>"
        ),
        lambda pkg: pkg.main_document_part,
        lambda pkg: pkg.main_document_part.contents.body.content[0].sdt_content.content,
    ),
    "w:customXml (block)": (
        "block",
        lambda: Probe(
            f'<w:customXml w:element="x">{p("before")}{ac("block", "CXB")}{p("after")}</w:customXml>'
        ),
        lambda pkg: pkg.main_document_part,
        lambda pkg: pkg.main_document_part.contents.body.content[0].content,
    ),
    "w:docPartBody": (
        "block",
        lambda: Probe(p("body")).part(
            "glossary/document.xml",
            "document.glossary",
            "glossaryDocument",
            "rIdG",
            "w:glossaryDocument",
            '<w:docParts><w:docPart><w:docPartPr><w:name w:val="n"/></w:docPartPr>'
            f"<w:docPartBody>{p('before')}{ac('block', 'GLOS')}{p('after')}</w:docPartBody></w:docPart></w:docParts>",
        ),
        lambda pkg: _part(pkg, GlossaryDocumentPart),
        lambda pkg: (
            _part(pkg, GlossaryDocumentPart)
            .contents.doc_parts.doc_part[0]
            .doc_part_body.content
        ),
    ),
    "w:tbl": (
        "row",
        lambda: Probe(tbl(tr("before") + ac("row", "TBL") + tr("after"))),
        lambda pkg: pkg.main_document_part,
        lambda pkg: pkg.main_document_part.contents.body.content[0].content,
    ),
    "w:sdtContent (row)": (
        "row",
        lambda: Probe(
            tbl(
                f"<w:sdt><w:sdtPr/><w:sdtContent>{tr('before')}{ac('row', 'SDTR')}{tr('after')}</w:sdtContent></w:sdt>"
            )
        ),
        lambda pkg: pkg.main_document_part,
        lambda pkg: (
            _first(pkg.main_document_part.contents, CTSdtRow).sdt_content.content
        ),
    ),
    "w:customXml (row)": (
        "row",
        lambda: Probe(
            tbl(
                f'<w:customXml w:element="x">{tr("before")}{ac("row", "CXR")}{tr("after")}</w:customXml>'
            )
        ),
        lambda pkg: pkg.main_document_part,
        lambda pkg: pkg.main_document_part.contents.body.content[0].content[0].content,
    ),
    "w:tr": (
        "cell",
        lambda: Probe(
            tbl(f"<w:tr>{tc('before')}{ac('cell', 'TR')}{tc('after')}</w:tr>")
        ),
        lambda pkg: pkg.main_document_part,
        lambda pkg: _first(pkg.main_document_part.contents, Tr).content,
    ),
    "w:sdtContent (cell)": (
        "cell",
        lambda: Probe(
            tbl(
                f"<w:tr><w:sdt><w:sdtPr/><w:sdtContent>{tc('before')}{ac('cell', 'SDTC')}{tc('after')}</w:sdtContent></w:sdt></w:tr>"
            )
        ),
        lambda pkg: pkg.main_document_part,
        lambda pkg: (
            _first(pkg.main_document_part.contents, CTSdtCell).sdt_content.content
        ),
    ),
    "w:customXml (cell)": (
        "cell",
        lambda: Probe(
            tbl(
                f'<w:tr><w:customXml w:element="x">{tc("before")}{ac("cell", "CXC")}{tc("after")}</w:customXml></w:tr>'
            )
        ),
        lambda pkg: pkg.main_document_part,
        lambda pkg: _first(pkg.main_document_part.contents, Tr).content[0].content,
    ),
}

_PROMOTED = {"block": P, "row": Tr, "cell": Tc}
_TAGS = {
    "w:body": "BODY",
    "w:hdr": "HDR",
    "w:ftr": "FTR",
    "w:tc": "TC",
    "w:footnote": "FN",
    "w:endnote": "EN",
    "w:comment": "CMT",
    "w:txbxContent": "TXBX",
    "w:sdtContent (block)": "SDTB",
    "w:customXml (block)": "CXB",
    "w:docPartBody": "GLOS",
    "w:tbl": "TBL",
    "w:sdtContent (row)": "SDTR",
    "w:customXml (row)": "CXR",
    "w:tr": "TR",
    "w:sdtContent (cell)": "SDTC",
    "w:customXml (cell)": "CXC",
}


def _saved_part(pkg: Any, part: Any) -> bytes:
    name = str(part.part_name).lstrip("/")
    return zipfile.ZipFile(io.BytesIO(pkg.save())).read(name)


# ---------------------------------------------------------------------------
# the probes, in both modes
# ---------------------------------------------------------------------------


def test_the_probes_are_docx4js():
    assert len(PROBES) == 19
    assert len(WITH_WORD_SAVE) == 14


@pytest.mark.parametrize("mode", MODES)
@pytest.mark.parametrize("stem", PROBES)
def test_no_probe_skips_anything(stem: str, mode: str) -> None:
    pkg = probe(stem, mode)
    for part in wml_parts(pkg):
        assert part.contents is not None
        assert part.skipped == [], (part.part_name, part.skipped[:3])


@pytest.mark.parametrize("mode", MODES)
@pytest.mark.parametrize("stem", PROBES)
def test_strict_loads_every_probe(stem: str, mode: str) -> None:
    pkg = load(
        PROBES_DIR / f"{stem}.docx",
        options=LoadOptions(strict=True, mc_block_level=mode),
    )
    for part in wml_parts(pkg):
        assert part.contents is not None


@pytest.mark.parametrize("stem", PROBES)
def test_preserve_saves_each_probe_canonically_identical(stem: str) -> None:
    pkg = probe(stem)
    source = (PROBES_DIR / f"{stem}.docx").read_bytes()
    saved = main_xml(pkg.save())
    assert normalize(saved) == normalize(main_xml(source))
    for token in ("<mc:AlternateContent", "<mc:Choice ", "<mc:Fallback"):
        assert count(saved, token) == count(main_xml(source), token), token
    assert pkg.main_document_part.resolved_alternate_content == []


def test_preserve_is_the_default():
    assert LoadOptions().mc_block_level == "preserve"
    pkg = load(PROBES_DIR / "mc-block-newest-first.docx")
    assert pkg.mc_block_level == "preserve"
    assert isinstance(pkg.main_document_part.contents.body.content[2], AlternateContent)


#: docx4j McBlockLevelProbesTest.paperSection78: Word 365's branch, original order and newest first.
PAPER_78 = {
    "mc-text": "w14",
    "mc-run": "w14",
    "mc-block": "w14",
    "mc-text-newest-first": "w16sdtfl",
    "mc-run-newest-first": "w16sdtfl",
    "mc-block-newest-first": "w16sdtfl",
}
_OTHERS = (
    "w15",
    "w16se",
    "w16cid",
    "w16cex",
    "w16sdtdh",
    "w16du",
    "wp14",
    "fallback",
    "TheEarthisround",
)


@pytest.mark.parametrize("mode", MODES)
@pytest.mark.parametrize(("stem", "branch"), PAPER_78.items())
def test_each_mode_reads_the_branch_word_365_shows(
    stem: str, branch: str, mode: str
) -> None:
    shown = text(probe(stem, mode)).rstrip()
    assert shown.endswith(branch), shown
    assert not any(other in shown for other in _OTHERS if other != branch), shown


def _description_differences(stem: str) -> list[str]:
    pkg = probe(stem, "word")
    mine = describe(load(io.BytesIO(pkg.save()))).split("\n")
    words = describe(load(PROBES_DIR / f"{stem}-word365.docx")).split("\n")
    remaining = list(words)
    out = []
    for line in mine:
        if line in remaining:
            remaining.remove(line)
        else:
            out.append("+ " + line)
    out.extend("- " + line for line in remaining)
    return out


#: Where a word-mode save is not Word's, and why (CR-012 section 9). Word 365 writes a text box
#: of a block-level element as a new run-level element, the wps box as the Choice and a VML copy
#: it regenerates as the Fallback (CR-051 section 3.5); this package writes the bare wps box, so
#: Word's description has the box twice. And Word recomputes the table grid of the cell probe.
EXPECTED_DIFFERENCES = {
    "mc-txbx-block": 2,
    "mc-txbx-sdt-block": 2,
    "mc-txbx-tc-block": 4,
}


@pytest.mark.parametrize("stem", WITH_WORD_SAVE)
def test_a_word_mode_save_describes_as_word_365s_save(stem: str) -> None:
    found = _description_differences(stem)
    assert len(found) == EXPECTED_DIFFERENCES.get(stem, 0), found
    for line in found:
        assert "txbx{" in line or "grid[" in line, line


@pytest.mark.parametrize("stem", WITH_WORD_SAVE)
def test_a_word_mode_save_reads_as_word_365s_save(stem: str) -> None:
    """docx4j McBlockLevelProbesTest.wordModeSaveReadsAsWordsSave: the text model, before and after the save."""
    pkg = probe(stem, "word")
    words = text(load(PROBES_DIR / f"{stem}-word365.docx"))
    assert text(load(io.BytesIO(pkg.save()))) == words
    assert text(pkg) == words


def test_the_block_level_probe():
    """docx4j McBlockLevelProbesTest.blockLevelProbe: Word's branch at the element's index, the record, mc:Ignorable."""
    pkg = probe("mc-block-newest-first", "word")
    part = pkg.main_document_part
    content = part.contents.body.content
    assert [type(item).__name__ for item in content] == ["P", "P", "P"]
    assert text_of(content[2]) == "w16sdtfl"
    (record,) = part.resolved_alternate_content
    assert isinstance(record, ResolvedAlternateContent)
    assert (record.container_name, record.level, record.index, record.promoted) == (
        "w:body",
        "block",
        2,
        1,
    )
    assert record.address == "w:document/w:body[0]"
    assert isinstance(record.branch, AlternateContentChoice)
    assert record.branch.requires == "w16sdtfl"
    assert len(record.element.choice) == 10
    assert record.element.fallback is not None
    assert part.contents.ignorable == "w16sdtfl"
    saved = main_xml(pkg.save())
    assert count(saved, "<mc:AlternateContent") == 0
    root = re.search(rb"<w:document[^>]*>", saved).group(0)
    assert b'mc:Ignorable="w16sdtfl"' in root
    assert b"xmlns:w16sdtfl=" in root
    assert (
        repr(record)
        == 'mc:AlternateContent at w:document/w:body[0] index 2: Choice Requires="w16sdtfl", 1 promoted'
    )


def test_the_no_fallback_probe_takes_its_first_understood_choice():
    """docx4j McBlockLevelProbesTest.noFallbackProbe: ten Choices and no Fallback, the first understood, as Word 365.

    CR-012 section 5 lists it as REQ-305's kept-whole case; it is not, since its first Choice
    (``w16sdtfl``) is understood (CR-012 section 9). REQ-305's case is the next test.
    """
    pkg = probe("mc-block-nofallback", "word")
    assert text_of(pkg.main_document_part.contents.body.content[2]) == "w16sdtfl"
    assert len(pkg.main_document_part.resolved_alternate_content) == 1


def test_the_mixed_probe_resolves_the_bodys_element_only():
    """docx4j McBlockLevelProbesTest.mixedProbeResolvesTheBodysElementOnly."""
    pkg = probe("mc-mixed", "word")
    assert len(pkg.main_document_part.resolved_alternate_content) == 1
    assert count(main_xml(pkg.save()), "<mc:AlternateContent") == 3
    shown = text(pkg)
    for marker in ("SDTRUN_W14", "HLINK_W14", "INP_W14", "BODY_W14"):
        assert marker in shown
    assert "FALLBACK" not in shown


@pytest.mark.parametrize(
    "stem", [s for s in PROBES if "newest-first" in s or s == "mc-block-nofallback"]
)
def test_mce_preprocess_takes_w16sdtfl_on_the_newest_first_probes(stem: str) -> None:
    """CR-012 Q8: over docx4j's published list, as Word 365; ``UNDERSTOOD`` took ``w16cid``."""
    pkg = load(PROBES_DIR / f"{stem}.docx", options=LoadOptions(mce_preprocess=True))
    shown = text(pkg)
    assert "w16sdtfl" in shown
    assert "w16cid" not in shown


# ---------------------------------------------------------------------------
# word mode: the splice, in every container (docx4j McBlockLevelLoadTest)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", CASES)
def test_word_mode_splices_every_container(name: str) -> None:
    kind, made, part_of, content_of = CASES[name]
    tag = _TAGS[name]
    pkg = made().load("word")
    part = part_of(pkg)
    content = content_of(pkg)
    assert len(content) == 3
    at = content[1]
    assert isinstance(at, _PROMOTED[kind]), type(at)
    assert f"{tag}_W14" in text_of(at)
    assert not any(isinstance(item, AlternateContent) for item in content)
    assert find(part.contents, AlternateContent, mce="all") == []
    assert at.parent is content[0].parent
    (record,) = part.resolved_alternate_content
    assert record.index == 1 and record.promoted == 1
    assert record.container_name == name.split(" ")[0]
    assert record.level == kind
    assert re.search(
        re.escape(record.container_name) + r"(\[\d+\])?$", record.address
    ), record.address
    assert record.branch.requires == "w14"
    assert len(record.element.choice) == 2 and record.element.fallback is not None
    assert f"{tag}_FB" in to_xml(record.element)
    assert "w14" in (part.contents.ignorable or "").split()
    saved = load(pkg.save(), options=LoadOptions(mc_block_level="word"))
    again = part_of(saved)
    assert f"{tag}_W14" in text_of(content_of(saved)[1])
    assert find(again.contents, AlternateContent, mce="all") == []
    assert again.resolved_alternate_content == []
    assert "skipped" not in repr(part.skipped) and part.skipped == []


@pytest.mark.parametrize("name", CASES)
def test_preserve_mode_keeps_every_container_whole(name: str) -> None:
    _kind, made, part_of, content_of = CASES[name]
    source = made()
    pkg = source.load("preserve")
    part = part_of(pkg)
    content = content_of(pkg)
    assert len(content) == 3
    assert isinstance(content[1], AlternateContent)
    assert part.resolved_alternate_content == []
    assert part.skipped == []
    data = source.parts[str(part.part_name).lstrip("/")].encode()
    assert normalize(_saved_part(pkg, part)) == normalize(data)


def test_word_mode_leaves_run_and_paragraph_level_whole():
    made = Probe(
        '<w:p><w:r><w:t>x</w:t></w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>C</w:t></w:r></mc:Choice>'
        "<mc:Fallback><w:r><w:t>F</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>"
    )
    pkg = made.load("word")
    paragraph = pkg.main_document_part.contents.body.content[0]
    assert isinstance(paragraph.content[1], AlternateContent)
    assert pkg.main_document_part.resolved_alternate_content == []


def test_req_305_an_element_with_no_understood_choice_and_no_fallback_stays_whole():
    """REQ-305: kept whole in word mode too; a reader takes its first Choice."""
    element = f'<mc:AlternateContent><mc:Choice Requires="wpi">{p("WPI")}</mc:Choice><mc:Choice Requires="w99">{p("W99")}</mc:Choice></mc:AlternateContent>'
    made = Probe(p("x") + element + p("y"))
    made.parts["word/document.xml"] = made.parts["word/document.xml"].replace(
        "<w:document ", '<w:document xmlns:w99="urn:example:w99" '
    )
    for mode in MODES:
        pkg = made.load(mode)
        content = pkg.main_document_part.contents.body.content
        assert len(content) == 3 and isinstance(content[1], AlternateContent), mode
        assert pkg.main_document_part.resolved_alternate_content == []
        saved = main_xml(pkg.save())
        assert b'Requires="wpi"' in saved and b'Requires="w99"' in saved
        assert "WPI" in text(pkg)


def test_an_element_with_no_branch_is_removed_in_word_mode_and_kept_in_preserve():
    made = Probe(p("x") + "<mc:AlternateContent/>" + p("y"))
    word = made.load("word")
    assert len(word.main_document_part.contents.body.content) == 2
    (record,) = word.main_document_part.resolved_alternate_content
    assert record.branch is None and record.index == 1 and record.promoted == 0
    assert "none (removed)" in repr(record)
    preserve = made.load("preserve")
    content = preserve.main_document_part.contents.body.content
    assert len(content) == 3 and isinstance(content[1], AlternateContent)


def test_nested_elements_resolve_outermost_first():
    """docx4j McBlockLevelLoadTest: an element in a cell inside the outer element's selected branch, and one in its discarded branch."""
    inner = (
        f'<mc:AlternateContent><mc:Choice Requires="w14">{p("INNER_W14")}</mc:Choice>'
        f"<mc:Fallback>{p('INNER_FB')}</mc:Fallback></mc:AlternateContent>"
    )
    discarded = (
        f'<mc:AlternateContent><mc:Choice Requires="w14">{p("DISCARDED_W14")}</mc:Choice>'
        f"<mc:Fallback>{p('DISCARDED_FB')}</mc:Fallback></mc:AlternateContent>"
    )
    outer = (
        f'<mc:AlternateContent><mc:Choice Requires="w14">{tbl("<w:tr><w:tc>" + p("cell") + inner + "</w:tc></w:tr>")}</mc:Choice>'
        f"<mc:Fallback>{p('OUTER_FB')}{discarded}</mc:Fallback></mc:AlternateContent>"
    )
    pkg = Probe(p("x") + outer + p("y")).load("word")
    part = pkg.main_document_part
    content = part.contents.body.content
    assert len(content) == 3
    cell = _first(part.contents, Tc)
    assert len(cell.content) == 2 and "INNER_W14" in text_of(cell.content[1])
    assert find(part.contents, AlternateContent, mce="all") == []
    records = part.resolved_alternate_content
    assert [r.container_name for r in records] == ["w:body", "w:tc"], (
        "the discarded branch's is not recorded"
    )
    assert "w:tbl[1]" in records[1].address
    outer_xml = to_xml(records[0].element)
    assert (
        "DISCARDED_W14" in outer_xml
        and "DISCARDED_FB" in outer_xml
        and "OUTER_FB" in outer_xml
    )


def test_an_element_kept_whole_keeps_the_elements_inside_it_whole():
    inner = f'<mc:AlternateContent><mc:Choice Requires="w14">{p("IN_W14")}</mc:Choice><mc:Fallback>{p("IN_FB")}</mc:Fallback></mc:AlternateContent>'
    outer = f'<mc:AlternateContent><mc:Choice Requires="wpi">{p("WPI")}{inner}</mc:Choice></mc:AlternateContent>'
    pkg = Probe(outer).load("word")
    assert isinstance(pkg.main_document_part.contents.body.content[0], AlternateContent)
    assert pkg.main_document_part.resolved_alternate_content == []
    assert count(main_xml(pkg.save()), "<mc:AlternateContent") == 2


def test_the_selected_choices_prefixes_join_mc_ignorable():
    """docx4j McBlockLevelLoadTest.selectedPrefixesJoinIgnorable: added to a list, set where the root had none, never twice."""
    element = f'<mc:AlternateContent><mc:Choice Requires="w15">{p("W15")}</mc:Choice><mc:Fallback>{p("FB")}</mc:Fallback></mc:AlternateContent>'
    listed = Probe(p("x") + element, ' mc:Ignorable="w14"').load("word")
    assert listed.main_document_part.contents.ignorable == "w14 w15"
    saved = main_xml(listed.save())
    root = re.search(rb"<w:document[^>]*>", saved).group(0)
    assert b'mc:Ignorable="w14 w15"' in root and b"xmlns:w15=" in root
    none = Probe(p("x") + element).load("word")
    assert none.main_document_part.contents.ignorable == "w15"
    already = Probe(p("x") + element, ' mc:Ignorable="w14 w15"').load("word")
    assert already.main_document_part.contents.ignorable == "w14 w15"


def test_a_part_without_the_element_pays_one_substring_test():
    data = (
        f"{_DECLARATION}<w:document {NS}><w:body>{p('plain')}</w:body></w:document>"
    ).encode()
    spliced, records = splice_word(data)
    assert spliced is data and records == []


def test_a_part_with_only_run_level_elements_is_parsed_from_its_own_bytes():
    pkg = probe("mc-run-newest-first", "word")
    assert pkg.main_document_part.resolved_alternate_content == []
    assert normalize(main_xml(pkg.save())) == normalize(
        main_xml((PROBES_DIR / "mc-run-newest-first.docx").read_bytes())
    )


# ---------------------------------------------------------------------------
# the mode: a load option and a package setting
# ---------------------------------------------------------------------------


def test_the_mode_is_checked():
    with pytest.raises(ValueError, match="mc_block_level"):
        LoadOptions(mc_block_level="Word")
    pkg = load(PROBES_DIR / "mc-block.docx")
    with pytest.raises(ValueError, match="mc_block_level"):
        pkg.mc_block_level = "resolve"


def test_the_package_setting_overrides_the_load_option_for_a_part_not_yet_parsed():
    pkg = load(
        PROBES_DIR / "mc-block-newest-first.docx",
        options=LoadOptions(mc_block_level="preserve"),
    )
    pkg.mc_block_level = "word"
    assert pkg.mc_block_level == "word"
    assert pkg.load_options.mc_block_level == "preserve", "the options are not changed"
    assert len(pkg.main_document_part.contents.body.content) == 3
    assert isinstance(pkg.main_document_part.contents.body.content[2], P)


def test_a_part_parsed_before_the_setting_keeps_its_tree():
    pkg = load(PROBES_DIR / "mc-block-newest-first.docx")
    tree = pkg.main_document_part.contents
    pkg.mc_block_level = "word"
    assert pkg.main_document_part.contents is tree
    assert isinstance(tree.body.content[2], AlternateContent)


def test_set_bytes_clears_the_record():
    pkg = probe("mc-block-newest-first", "word")
    part = pkg.main_document_part
    assert part.resolved_alternate_content
    part.set_bytes(
        (PROBES_DIR / "mc-block.docx").read_bytes()
        and main_xml((PROBES_DIR / "mc-block.docx").read_bytes())
    )
    assert part.resolved_alternate_content == []
    assert part.contents is not None
    assert len(part.resolved_alternate_content) == 1


def test_word_mode_and_mce_preprocess_together():
    pkg = load(
        PROBES_DIR / "mc-mixed.docx",
        options=LoadOptions(mc_block_level="word", mce_preprocess=True),
    )
    part = pkg.main_document_part
    assert part.contents is not None
    assert len(part.resolved_alternate_content) == 1, (
        "the block-level element is recorded"
    )
    assert find(part.contents, AlternateContent, mce="all") == []
    assert "BODY_W14" in text(pkg)


# ---------------------------------------------------------------------------
# a w:sdt straight in a branch: typed as its container's kind (the fork's stage 11)
# ---------------------------------------------------------------------------

#: CR-012 section 3.1's fragment: a run-level control straight in a paragraph's mc:Choice.
SECTION_31 = (
    f"<w:p {NS}><w:r><w:t>x</w:t></w:r><mc:AlternateContent>"
    '<mc:Choice Requires="w14"><w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr><w:sdtContent>'
    "<w:r><w:t>SDTRUN_W14</w:t></w:r></w:sdtContent></w:sdt></mc:Choice>"
    "<mc:Fallback><w:r><w:t>FALLBACK</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>"
)


def test_wml_takes_section_3_1s_fragment():
    """It raised FragmentError, naming the w:r skipped in CT_SdtContentBlock."""
    paragraph = wml(SECTION_31)
    (sdt,) = paragraph.content[1].choice[0].any_element
    assert type(sdt) is SdtRun
    assert text_of(sdt) == "SDTRUN_W14"
    assert normalize(to_xml(paragraph).encode()) == normalize(SECTION_31.encode())


def _branch_sdt(container: str, inner: str) -> str:
    return (
        f'<mc:AlternateContent><mc:Choice Requires="w14"><w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr>'
        f"<w:sdtContent>{inner}</w:sdtContent></w:sdt></mc:Choice>"
        f"<mc:Fallback>{p('FB') if container == 'body' else ''}</mc:Fallback></mc:AlternateContent>"
    )


SDT_CASES = {
    # level: (the body, the class, where the element is)
    "paragraph": (
        f"<w:p>{_branch_sdt('p', '<w:r><w:t>SDT_IN_P</w:t></w:r>')}</w:p>",
        SdtRun,
        lambda body: body.content[0].content[0],
    ),
    "block": (
        _branch_sdt("body", p("SDT_IN_BODY")),
        SdtBlock,
        lambda body: body.content[0],
    ),
    "row": (
        tbl(tr("r") + _branch_sdt("tbl", tr("SDT_IN_TBL"))),
        CTSdtRow,
        lambda body: body.content[0].content[1],
    ),
    "cell": (
        tbl(f"<w:tr>{tc('c')}{_branch_sdt('tr', tc('SDT_IN_TR'))}</w:tr>"),
        CTSdtCell,
        lambda body: body.content[0].content[0].content[1],
    ),
}


@pytest.mark.parametrize("level", SDT_CASES)
def test_a_sdt_straight_in_a_branch_is_its_containers_kind(level: str) -> None:
    body_xml, cls, element_of = SDT_CASES[level]
    made = Probe(body_xml)
    pkg = made.load("preserve", strict=True)
    part = pkg.main_document_part
    element = element_of(part.contents.body)
    assert isinstance(element, AlternateContent)
    (sdt,) = element.choice[0].any_element
    assert type(sdt) is cls
    assert part.skipped == []
    assert normalize(main_xml(pkg.save())) == normalize(
        made.parts["word/document.xml"].encode()
    )


@pytest.mark.parametrize("level", ["block", "row", "cell"])
def test_a_sdt_promoted_in_word_mode_is_its_containers_kind(level: str) -> None:
    body_xml, cls, element_of = SDT_CASES[level]
    pkg = Probe(body_xml).load("word")
    promoted = element_of(pkg.main_document_part.contents.body)
    assert type(promoted) is cls
    (record,) = pkg.main_document_part.resolved_alternate_content
    (kept,) = record.element.choice[0].any_element
    assert type(kept) is cls, "the record's own parse types it by the level"


# ---------------------------------------------------------------------------
# the readers' audit (section 4.1 item 5), in preserve mode
# ---------------------------------------------------------------------------


def test_traversal_reads_words_branch():
    body = probe("mc-block-newest-first").main_document_part.contents.body
    assert text_of(body) == "Block level test\n\nw16sdtfl"
    assert [text_of(p_) for p_ in find(body, P)] == ["Block level test", "", "w16sdtfl"]
    assert len(find(body, P, mce="all")) == 13


def test_the_body_reads_words_branch_at_block_level():
    pkg = probe("mc-sdt-block-newest-first")
    paragraphs = pkg.body.paragraphs
    assert [(p_.text, p_.address) for p_ in paragraphs] == [
        ("Block sdt test", "body/0"),
        ("", "body/1"),
        ("w16sdtfl", "body/2/0/0"),
    ]
    found = pkg.body.element_at("body/2/0/0")
    assert found.text == "w16sdtfl"


@pytest.mark.parametrize(
    "stem",
    [
        "mc-row-newest-first",
        "mc-sdt-row-newest-first",
        "mc-cell-newest-first",
        "mc-sdt-cell-newest-first",
    ],
)
def test_the_body_passes_row_and_cell_level_by(stem: str) -> None:
    """As docx4j's CR-051 leaves it until its phase 2 (CR-012 section 9): no view, no error."""
    pkg = probe(stem)
    texts = [p_.text for p_ in pkg.body.paragraphs]
    assert "w16sdtfl" not in texts and "row" not in texts
    (table,) = pkg.body.tables
    assert [[cell.body.text for cell in row.cells] for row in table.rows] in (
        [["I prefer", "below"]],
        [["I prefer"]],
    )
    pkg.outline()
    pkg.to_markdown()
    pkg.describe()


@pytest.mark.parametrize(
    ("stem", "shown"),
    [
        ("mc-row-newest-first", "Row level test\nI prefer\tbelow\n\n"),
        ("mc-sdt-row-newest-first", "Row sdt test\nI prefer\tbelow\n\n"),
        ("mc-cell-newest-first", "Cell level test\nI prefer\n\n"),
        ("mc-sdt-cell-newest-first", "Cell sdt test\nI prefer\n\n"),
    ],
)
def test_the_text_model_passes_row_and_cell_level_by(stem: str, shown: str) -> None:
    """As docx4j's CR-051 phase 1 leaves it; its phase 2 (9db1a3cbd) reads them, a follow-up here."""
    assert text(probe(stem)) == shown
    assert text(probe(stem, "word")) != shown, (
        "word mode reads Word's branch, spliced in on load"
    )


def test_an_edit_inside_a_kept_branch_is_refused():
    pkg = probe("mc-block-newest-first")
    branch_paragraph = pkg.body.paragraphs[2]
    for edit in (
        lambda: branch_paragraph.insert_text("X"),
        lambda: setattr(branch_paragraph, "text", "X"),
        lambda: branch_paragraph.insert_paragraph("X", location="After"),
        lambda: setattr(branch_paragraph.font, "bold", True),
        lambda: pkg.body.element_at(branch_paragraph.address).insert_text("X"),
    ):
        with pytest.raises(AddressError) as caught:
            edit()
        assert caught.value.code == "address.read_only"
    assert normalize(main_xml(pkg.save())) == normalize(
        main_xml((PROBES_DIR / "mc-block-newest-first.docx").read_bytes())
    )
    pkg.body.paragraphs[0].insert_text(" ok")
    assert pkg.body.paragraphs[0].text == "Block level test ok"


def test_an_edit_there_in_word_mode_is_made():
    pkg = probe("mc-block-newest-first", "word")
    pkg.body.paragraphs[2].insert_text("X")
    assert pkg.body.paragraphs[2].text == "w16sdtflX"


def test_the_outline_reads_words_branch():
    pkg = probe("mc-block-newest-first")
    outline = pkg.outline()
    entries = [(e.address, e.kind, e.text) for e in outline.entries]
    assert entries[:2] == [
        ("body/0", "paragraph", "Block level test"),
        ("body/1", "paragraph", ""),
    ]
    (element,) = [e for e in outline.entries if e.address == "body/2"]
    assert [(c.address, c.kind, c.text) for c in element.children] == [
        ("body/2/0", "paragraph", "w16sdtfl")
    ]
    assert outline.stats.paragraphs == 3


def test_describe_passes_the_element_by():
    description = probe("mc-block-newest-first").describe()
    assert description.skipped == 0


@pytest.mark.parametrize("mode", MODES)
def test_the_text_model_and_markdown_read_words_branch(mode: str) -> None:
    """Closes CR-011 section 14.7 item 8 (docx4j MarkdownAlternateContentTest, both modes)."""
    pkg = probe("mc-block-newest-first", mode)
    assert text(pkg).endswith("w16sdtfl\n")
    assert pkg.to_markdown().endswith("w16sdtfl\n")


def test_the_text_model_warns_where_the_branches_differ():
    from docx4j_py.model.text import TextOptions

    result = extract(probe("mc-block-newest-first"), TextOptions())
    warnings = [w for w in result.warnings if w.kind == "mce.branches_differ"]
    assert [(w.address, w.branch) for w in warnings] == [
        ("body/2", 'mc:Choice Requires="w16sdtfl"')
    ]


def test_the_python_docx_facade_passes_the_element_by():
    from docx4j_py import docx

    document = docx.Document(str(PROBES_DIR / "mc-block-newest-first.docx"))
    assert [p_.text for p_ in document.paragraphs] == ["Block level test", ""]
    document.paragraphs[0].add_run(" ok")
    assert document.paragraphs[0].text == "Block level test ok"


def test_the_revisions_engine_passes_row_and_cell_level_by():
    from docx4j_py.model import revisions

    for stem in (
        "mc-row-newest-first",
        "mc-cell-newest-first",
        "mc-sdt-row-newest-first",
    ):
        pkg = probe(stem)
        assert revisions.tracked_changes(pkg) == []
        assert revisions.accept_all(pkg) == 0


def test_a_part_that_never_held_the_element_is_byte_for_byte_in_either_mode():
    for mode in MODES:
        pkg = load(
            PROBES_DIR / "mc-block-newest-first.docx",
            options=LoadOptions(mc_block_level=mode),
        )
        saved = zipfile.ZipFile(io.BytesIO(pkg.save()))
        source = zipfile.ZipFile(PROBES_DIR / "mc-block-newest-first.docx")
        assert saved.read("word/document.xml") == source.read("word/document.xml"), (
            "never unmarshalled"
        )


def test_the_record_survives_an_edit():
    pkg = probe("mc-block-newest-first", "word")
    (record,) = pkg.main_document_part.resolved_alternate_content
    pkg.body.paragraphs[0].insert_text("!")
    assert pkg.main_document_part.resolved_alternate_content == [record]


def test_a_fragment_keeps_a_block_level_element_whole():
    """Content parsed outside a part never passes the splice (CR-051 section 3.4)."""
    body = wml(f"<w:body {NS}>{p('x')}{ac('block', 'FR')}</w:body>")
    assert isinstance(body.content[1], AlternateContent)
    with pytest.raises(FragmentError):
        wml(f"<w:body {NS}><w:bogus/></w:body>")


def test_a_corrupt_part_still_raises_in_word_mode():
    made = Probe(p("x") + ac("block", "B"))
    made.parts["word/document.xml"] = made.parts["word/document.xml"].replace(
        "</w:body>", ""
    )
    pkg = made.load("word")
    with pytest.raises((Docx4JException, etree.XMLSyntaxError)):
        assert pkg.main_document_part.contents is not None


def _branch_warnings(pkg: Any) -> list[tuple[str | None, str | None]]:
    from docx4j_py.model.text import TextOptions

    options = TextOptions().include("headers", "footers", "comments", "glossary")
    result = extract(pkg, options)
    return sorted(
        (w.address, w.branch)
        for w in result.warnings
        if w.kind == "mce.branches_differ"
    )


@pytest.mark.parametrize(
    "stem", [s for s in PROBES if "row" not in s and "cell" not in s]
)
def test_the_text_models_warnings_are_the_same_in_both_modes_on_the_probes(
    stem: str,
) -> None:
    """docx4j McBlockLevelReadersTest: in word mode the warning is raised from the record, where preserve mode raises it."""
    assert _branch_warnings(probe(stem, "word")) == _branch_warnings(
        probe(stem, "preserve")
    )


@pytest.mark.parametrize(
    "name", [name for name, case in CASES.items() if case[0] == "block"]
)
def test_the_text_models_warnings_are_the_same_in_both_modes_in_every_block_container(
    name: str,
) -> None:
    _kind, made, _part_of, _content_of = CASES[name]
    preserve = _branch_warnings(made().load("preserve"))
    assert _branch_warnings(made().load("word")) == preserve
    assert len(preserve) == 1, preserve
    assert preserve[0][1] == 'mc:Choice Requires="w14"'


def test_a_custom_xml_element_straight_in_a_branch():
    """docx4j McBlockLevelLoadTest's custom XML in a branch: typed once word mode promotes it; whole in preserve."""
    element = (
        '<mc:AlternateContent><mc:Choice Requires="w14"><w:customXml w:element="x">'
        f"{p('CX_W14')}</w:customXml></mc:Choice><mc:Fallback>{p('CX_FB')}</mc:Fallback></mc:AlternateContent>"
    )
    made = Probe(p("before") + element)
    word = made.load("word")
    promoted = word.main_document_part.contents.body.content[1]
    assert type(promoted) is CTCustomXmlBlock
    assert text_of(promoted) == "CX_W14"
    saved = main_xml(word.save())
    assert b'<w:customXml w:element="x">' in saved and b"CX_W14" in saved
    preserve = made.load("preserve")
    assert isinstance(
        preserve.main_document_part.contents.body.content[1], AlternateContent
    )
    assert preserve.main_document_part.skipped == []
    assert normalize(main_xml(preserve.save())) == normalize(
        made.parts["word/document.xml"].encode()
    )
