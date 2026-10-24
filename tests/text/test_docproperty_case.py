"""DOCPROPERTY names matched in any case, as Word 365 matches them (docx4j ``ea5461ee9``, #244).

docx4j ``VERSION_17_3_5``: ``DocPropertyResolver`` takes a custom property's name
exactly first, else in any case, and the built-in names in any case; and
``Keywords`` is a core property (``f15bce627``, #277). The text model computes
``DOCPROPERTY`` under ``fields="update"`` (CR-011, REQ-506), so these are its
results. The probe and Word's save of it are docx4j's (``tests/fixtures/docprop-case/``).
"""

from __future__ import annotations

import hashlib
import re
from pathlib import Path

import pytest
from mini_docx import MiniDocx

from docx4j_py import load
from docx4j_py.model.fields import DocPropertyResolver
from docx4j_py.model.fields.formatting import FieldValueError
from docx4j_py.model.text import extract, plain_text

PROBE = Path(__file__).parent.parent / "fixtures" / "docprop-case"


def _field(instr: str, cached: str) -> str:
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
        f'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{cached}</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def _package(*instructions: str):
    """A custom property Client = Acme, the title "Probe title" and the keywords "alpha, beta"; a paragraph per field."""
    return (
        MiniDocx.body("".join(f"<w:p>{_field(instr, 'old')}</w:p>" for instr in instructions))
        .package_part(
            "docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "rIdCore",
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Probe title</dc:title>'
            "<cp:keywords>alpha, beta</cp:keywords></cp:coreProperties>",
        )
        .package_part(
            "docProps/custom.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
            "rIdCustom",
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Client"><vt:lpwstr>Acme</vt:lpwstr>'
            "</property></Properties>",
        )
        .load()
    )


def _results(pkg, **options) -> dict[str, str]:
    """Each marker's result, ``K<n> DOCPROPERTY ...: <result> (end)``."""
    out = {}
    for line in plain_text(extract(pkg, **options)).splitlines():
        found = re.match(r"(K\d) DOCPROPERTY [^:]*: (.*) \(end\)$", line)
        if found:
            out[found.group(1)] = found.group(2)
    return out


def test_the_probe_files_are_docx4js():
    sums = dict(reversed(line.split()) for line in (PROBE / "SHA1SUMS").read_text().splitlines())
    for name, digest in sums.items():
        assert hashlib.sha1((PROBE / name).read_bytes()).hexdigest() == digest


def test_the_probe_updated_gives_words_results_but_for_a_missing_property():
    word = _results(load(PROBE / "docprop-case-word365.docx"))
    assert word == {
        **{k: "Acme Ltd" for k in ("K1", "K2", "K3", "K4")},
        "K5": "The Title",
        "K6": "The Title",
        "K7": "Error! Unknown document property name.",
    }
    probe = load(PROBE / "docprop-case.docx")
    assert set(_results(probe).values()) == {"OLD"}  # cached results by default
    # Word's every result, but K7's: no such property, so the field keeps its cached result (docx4j keeps the old one)
    assert _results(probe, fields="update") == {**word, "K7": "OLD"}


def test_a_custom_name_is_matched_exactly_first_then_in_any_case():
    resolver = DocPropertyResolver(_package())
    assert [resolver.value(name) for name in ("Client", "client", "CLIENT", "cLiEnT")] == ["Acme"] * 4
    assert resolver.value("title") == resolver.value("Title") == "Probe title"
    with pytest.raises(FieldValueError):
        resolver.value("Clients")


def test_keywords_is_a_core_property():
    resolver = DocPropertyResolver(_package())
    assert resolver.value("Keywords") == resolver.value("keywords") == "alpha, beta"


def test_the_text_model_reads_each_name_in_any_case():
    """docx4j's own test of #244 (``DocInfoFieldsTest.aPropertyNameInAnyCase``), through the text model."""
    instructions = ("DOCPROPERTY client", "DOCPROPERTY CLIENT", 'DOCPROPERTY "cLiEnT"', "DOCPROPERTY title", "DOCPROPERTY keywords")
    pkg = _package(*instructions)
    assert plain_text(extract(pkg)).splitlines() == ["old"] * 5
    assert plain_text(extract(pkg, fields="update")).splitlines() == ["Acme", "Acme", "Acme", "Probe title", "alpha, beta"]
