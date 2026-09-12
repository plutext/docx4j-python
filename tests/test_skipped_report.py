"""The fork's skipped-content report, CR-001 section 7 and section 9.

Lenient parsing is what real documents need -- a `.docx` may carry markup no
schema in `schemas/` knows -- but lenient and silent is not acceptable when
round-trip fidelity is the promise. `ParserConfig(skipped_report=True)`
collects what was dropped; these tests prove it catches an unknown element and
an unknown attribute, that the report is per parse, and that the round-trip
harness fails on a part that has any.

    .venv-fork/bin/python -m pytest tests/test_skipped_report.py -q
"""

from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_xsdata.formats.dataclass.parsers import XmlParser
from docx4j_xsdata.formats.dataclass.parsers.config import ParserConfig
from docx4j_xsdata.formats.dataclass.parsers.skipped import SkippedReport

from docx4j_py.wml import Document

FIXTURES = ROOT / "tests" / "fixtures"
UNKNOWN = FIXTURES / "unknown_content.xml"
WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ZZ = "urn:docx4j-python:test:unknown"

CLEAN = (
    f'<w:document xmlns:w="{WML}"><w:body><w:p><w:r><w:t>known</w:t>'
    "</w:r></w:p></w:body></w:document>"
).encode()


def lenient_parser() -> XmlParser:
    config = ParserConfig(
        fail_on_unknown_properties=False,
        fail_on_unknown_attributes=False,
        fail_on_converter_warnings=False,
        skipped_report=SkippedReport(log=False),
    )
    return XmlParser(config=config)


@pytest.fixture
def parser() -> XmlParser:
    return lenient_parser()


def test_the_document_still_parses(parser: XmlParser) -> None:
    doc = parser.from_bytes(UNKNOWN.read_bytes(), Document)

    # Lenient means the known content survives; the two intruders are gone.
    assert isinstance(doc, Document)
    assert len(doc.body.content) == 2
    assert doc.body.content[0].content[0].content[0].value == "known"


def test_an_unknown_element_is_reported(parser: XmlParser) -> None:
    parser.from_bytes(UNKNOWN.read_bytes(), Document)

    elements = [n for n in parser.skipped if n.kind == "element"]
    assert len(elements) == 1
    node = elements[0]
    assert node.qname == f"{{{ZZ}}}widget"
    assert node.parent_qname == f"{{{WML}}}body"
    assert node.parent_class == "Body"
    # Only the outermost skipped element: zz:inner went with its parent.
    assert node.path.endswith(f"{{{WML}}}body/{{{ZZ}}}widget")


def test_an_unknown_attribute_is_reported(parser: XmlParser) -> None:
    parser.from_bytes(UNKNOWN.read_bytes(), Document)

    attributes = [n for n in parser.skipped if n.kind == "attribute"]
    assert len(attributes) == 1
    node = attributes[0]
    assert node.qname == f"{{{ZZ}}}flavour"
    assert node.parent_class == "P"
    assert node.path.endswith(f"@{{{ZZ}}}flavour")


def test_both_are_reported_together(parser: XmlParser) -> None:
    parser.from_bytes(UNKNOWN.read_bytes(), Document)

    assert len(parser.skipped) == 2
    assert {n.kind for n in parser.skipped} == {"element", "attribute"}


def test_the_report_is_per_parse(parser: XmlParser) -> None:
    data = UNKNOWN.read_bytes()

    parser.from_bytes(data, Document)
    assert len(parser.skipped) == 2

    # Parsing the same document again does not accumulate.
    parser.from_bytes(data, Document)
    assert len(parser.skipped) == 2

    # And a document with nothing unknown in it clears the report.
    parser.from_bytes(CLEAN, Document)
    assert len(parser.skipped) == 0


def test_a_real_part_skips_nothing() -> None:
    """The corpus is clean now that the anyAttribute patch is in."""
    parser = lenient_parser()
    with zipfile.ZipFile(ROOT / "samples" / "sample-docx.docx") as z:
        parser.from_bytes(z.read("word/document.xml"), Document)

    assert list(parser.skipped) == []


def test_the_harness_fails_on_skipped_content(tmp_path: Path) -> None:
    """scripts/roundtrip.py exits non-zero and names what was dropped."""
    sample = tmp_path / "skipped.docx"
    with zipfile.ZipFile(ROOT / "samples" / "sample-docx.docx") as src:
        with zipfile.ZipFile(sample, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "word/document.xml":
                    data = UNKNOWN.read_bytes()
                dst.writestr(item, data)

    result = subprocess.run(
        [
            sys.executable,
            str(ROOT / "scripts" / "roundtrip.py"),
            "--models-module",
            "docx4j_py.wml",
            "--runtime",
            "docx4j_xsdata",
            "--samples",
            str(tmp_path),
            "--out",
            str(tmp_path / "out"),
            "--repeat",
            "1",
        ],
        capture_output=True,
        text=True,
    )

    assert result.returncode != 0, result.stdout
    assert "SKIPPED element" in result.stdout
    assert "SKIPPED attribute" in result.stdout
    assert "ROUND TRIP NOT CLEAN" in result.stdout
    assert "2 skipped node(s)" in result.stdout
