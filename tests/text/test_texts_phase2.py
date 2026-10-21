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
from docx4j_py.model.text import ImageDescription, TextResult, TextWarning, extract, plain_text

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


# ---------------------------------------------------------------------------
# images and objects (REQ-514, REQ-512)
# ---------------------------------------------------------------------------


def runs(result, kind):
    return [r for b in result.blocks for r in b.runs if r.kind == kind]


def test_image_descriptions(probe):
    """REQ-514, REQ-707: the alt text, else the name; nothing for a decorative picture; the title never."""
    value = text(probe, alt_text=True)
    assert between(value, "Alt text and title:", "Case 34") == (
        "Alt text and title: [image: IMG33_ALT_TEXT]\nDecorative: \nSVG, its PNG fallback in a:blip: [image: IMG33_SVG_ALT]\n"
    )
    assert "IMG33_TITLE" not in value
    assert "[image" not in text(probe)
    images = runs(extract(probe), "IMAGE")
    decorative = [r for r in images if r.image.decorative]
    assert len(decorative) == 1
    assert decorative[0].description is None
    assert decorative[0].image.descr == "IMG33_DECORATIVE_DESCR"
    assert all(r.description.provenance == "AUTHOR" for r in images if r.description is not None)


def test_a_hosts_describer_is_marked_generated(probe):
    def captions(image):
        return None if image.decorative else ImageDescription.generated(f"a picture named {image.name}")

    assert "Alt text and title: [image, generated: a picture named Picture 14]" in text(
        probe, alt_text=True, image_describer=captions
    )


def test_a_vml_pictures_alt_and_no_decorative_block_in_the_probe(probe):
    vml = MiniDocx.body(
        '<w:p><w:r><w:pict><v:shape style="width:10pt;height:10pt" alt="VML alt"><v:imagedata r:id="rIdNone"/>'
        "</v:shape></w:pict></w:r></w:p>"
    ).load()
    assert text(vml, alt_text=True) == "[image: VML alt]\n"
    assert not any(b.decorative for b in extract(probe).blocks)


def test_a_paragraph_of_decorative_pictures_is_a_decorative_block():
    picture = (
        '<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="P" descr="D">'
        '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">'
        '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>'
        "</a:ext></a:extLst></wp:docPr><a:graphic>"
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
        '<pic:cNvPr id="0" name="P"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdNone"/>'
        "</pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )
    block = extract(MiniDocx.body(f"<w:p>{picture}</w:p>").load()).blocks[0]
    assert block.decorative and block.to_dict()["decorative"]
    assert block.runs[0].image.part is None  # the relationship names nothing


def test_a_describer_that_fails_raises_image_not_described(probe):
    def failing(image):
        raise RuntimeError("the describing service is down")

    r = extract(probe, alt_text=True, image_describer=failing)
    found = [w for w in r.warnings if w.kind == TextWarning.IMAGE_NOT_DESCRIBED]
    assert found and all(w.selector.startswith("/word/media/") for w in found)
    assert "[image" not in plain_text(r)


def test_objects_are_object_runs_with_their_progids(probe):
    assert [r.prog_id for r in runs(extract(probe), "OBJECT")] == [
        "Excel.Sheet.12",
        "Word.Document.12",
        "Package",
        None,
        "PowerPoint.Show.12",
    ]


def test_legacy_wordart_is_a_one_paragraph_text_box(probe):
    r = extract(probe)
    word_art = [b for b in r.blocks if b.story == "TEXT_BOX" and b.text == "WA25_TEXTPATH"]
    assert len(word_art) == 1 and word_art[0].address.endswith("/textbox:0/0")
    assert "WA25_TEXTPATH" not in text(probe, text_boxes=False)
