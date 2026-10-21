"""The text model's other content, ``docx4j_py.model.text``. CR-011 Phase C (section 13).

docx4j's ``TextsPhase2Test`` at ``198d015a3``, its strings kept: what the
suite's token rows do not pin on their own, built on the probes and on small
documents (``mini_docx.MiniDocx``, docx4j's ``MiniDocx``). Workbooks,
presentations and OLE 1.0 packages are deferred (CR-011 Q7), and their tests
are not ported.
"""

from __future__ import annotations

import pytest
from conformance import PHASE0
from mini_docx import MiniDocx

from docx4j_py import load
from docx4j_py.model.text import TextResult, extract, plain_text

PROBE = PHASE0 / "extraction-probe.docx"
EXTERNAL = PHASE0 / "extraction-probe-external.docm"


@pytest.fixture(scope="module")
def probe():
    return load(PROBE)


@pytest.fixture(scope="module")
def external():
    return load(EXTERNAL)


def text(package, **options) -> str:
    return plain_text(extract(package, **options))


def between(value: str, start: str, end: str) -> str:
    """The text from `start` up to `end`, `start` included."""
    i = value.index(start)
    return value[i : value.index(end, i + len(start))]


def p(value: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{value}</w:t></w:r></w:p>'


# ---------------------------------------------------------------------------
# metadata (REQ-515)
# ---------------------------------------------------------------------------


def test_metadata_always_in_the_result_and_in_the_text_only_under_metadata(probe, external):
    """REQ-515, REQ-801: the properties always in the metadata; in the text only under metadata: true, before the body."""
    r = extract(external)
    assert r.metadata[TextResult.FIELDS_UPDATE_ON_OPEN] is True
    assert r.metadata[TextResult.APP]["Company"] == "APP52_COMPANY"
    assert r.metadata[TextResult.CUSTOM]["ProbeCustom"] == "CPR52_CUSTOM"
    assert "APP52_COMPANY" not in plain_text(r)
    assert text(external, metadata=True).startswith(
        "[application properties: Application: Microsoft Office Word; Company: APP52_COMPANY]\n"
        "[custom properties: ProbeCustom: CPR52_CUSTOM]\nCase 50"
    )
    assert text(probe, metadata=True).startswith("[core properties: title: PRP38_TITLE; subject: PRP38_SUBJECT]\n")


def test_the_statistics_are_left_out():
    stats = (
        MiniDocx.body(p("Body."))
        .package_part(
            "docProps/app.xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
            "rIdApp",
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
            "<Template>Normal.dotm</Template><Pages>7</Pages><Words>1234</Words><Characters>5678</Characters>"
            "</Properties>",
        )
        .load()
    )
    assert extract(stats).metadata[TextResult.APP] == {"Template": "Normal.dotm"}


def test_every_core_property_in_the_schemas_order_and_custom_ones_typed():
    core = (
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<cp:revision>3</cp:revision><dc:creator>C</dc:creator><dc:title>T</dc:title><cp:keywords>K1 K2</cp:keywords>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-01T00:00:00Z</dcterms:created>'
        "<dc:description></dc:description></cp:coreProperties>"
    )
    custom = (
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Count"><vt:i4>7</vt:i4></property>'
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Done"><vt:bool>true</vt:bool></property>'
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="When">'
        "<vt:filetime>2026-01-02T03:04:05Z</vt:filetime></property>"
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Blob"><vt:clsid>{00000000-0000-0000-0000-000000000000}</vt:clsid></property>'
        "</Properties>"
    )
    pkg = (
        MiniDocx.body(p("Body."))
        .package_part(
            "docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "rIdCore",
            core,
        )
        .package_part(
            "docProps/custom.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
            "rIdCustom",
            custom,
        )
        .load()
    )
    r = extract(pkg, metadata=True)
    assert list(r.metadata[TextResult.CORE].items()) == [
        ("title", "T"),
        ("creator", "C"),
        ("keywords", "K1 K2"),
        ("revision", "3"),
        ("created", "2026-10-01T00:00:00Z"),
    ]
    assert r.metadata[TextResult.CUSTOM] == {"Count": 7, "Done": True, "When": "2026-01-02T03:04:05Z"}
    assert plain_text(r).startswith(
        "[core properties: title: T; creator: C; keywords: K1 K2; revision: 3; created: 2026-10-01T00:00:00Z]\n"
        "[custom properties: Count: 7; Done: true; When: 2026-01-02T03:04:05Z]\nBody.\n"
    )
    assert r.to_dict()["metadata"]["custom"]["Done"] is True
