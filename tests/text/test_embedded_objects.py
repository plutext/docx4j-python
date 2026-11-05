"""An object run's object, the embedded-object filter and a chunk warning's part: docx4j ``EmbeddedObjectsTest``. CR-027 Phase C.

docx4j's ``EmbeddedObjectsTest`` at ``docx4j-17.3.5`` (CR-070, ``1560ab379``;
WD 0.47 REQ-512, REQ-522, REQ-513, REQ-801), its nine tests:

- ``probeObjectsDescribed``, ``packageNames``, ``linkedObject``, ``controls``:
  as docx4j's, the OLE 1.0 packages built by ``ole_writer`` where docx4j's
  ``ole10`` uses POI;
- ``noFilterReadsEvery``, ``filterDeclinesOneOfTwo``, ``failingFilterDeclines``:
  **over embedded OOXML documents** (section 5's named departure), since
  docx4j's read ``NOTES_TEXT``, ``OTHER_TEXT`` and ``THIRD_TEXT`` out of OLE
  1.0 packages, whose content this package defers (CR-011 Q7), with docx4j's
  addresses and declines;
- ``filterOverTheProbe``: docx4j's four asks and ``EMB40_DOCX_TEXT``; the OLE
  package and the presentation it admits give ``content.not_read`` here, where
  docx4j expects ``OLE41_NATIVE_TEXT`` and ``EMB53_SLIDE_ONE``;
- ``chunkWarningNamesThePart``: as docx4j's.

Then this port's own: the same filter over OLE 1.0 packages (asked with
docx4j's addresses, each it admits ``content.not_read``); the block model's
JSON (section 4.3); a wildcard ``w:object``, which a kept branch holds; the
option itself; and a reading that leaves the parts as they were (REQ-204).
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from mini_docx import MiniDocx
from ole_writer import ole10

from docx4j_py import load
from docx4j_py.model.text import TextObject, TextOptionError, TextOptions, TextWarning, extract, plain_text
from docx4j_py.openpackaging.parts.binary_part import BinaryPart, EmbeddedPackagePart

ROOT = Path(__file__).resolve().parent.parent.parent
PROBE = ROOT / "tests" / "fixtures" / "text-extraction" / "phase0" / "extraction-probe.docx"

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
OLE_OBJECT = "application/vnd.openxmlformats-officedocument.oleObject"
ACTIVEX = "application/vnd.ms-office.activeX+xml"
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


@pytest.fixture(scope="module")
def probe():
    return load(PROBE)


def objects(result) -> list:
    return [r for b in result.blocks for r in b.runs if r.kind == "OBJECT"]


def _object(ole_type: str, prog_id: str, rel_id: str) -> str:
    return (
        f'<w:r><w:object><v:shape id="_x0000_{rel_id}" style="width:64pt;height:32pt"/>'
        f'<o:OLEObject Type="{ole_type}" ProgID="{prog_id}" ShapeID="_x0000_{rel_id}" r:id="{rel_id}"/>'
        "</w:object></w:r>"
    )


def _body(prog_id: str) -> str:
    """Two objects in one paragraph, and a third in the next."""
    return (
        f"<w:p><w:r><w:t>First</w:t></w:r>{_object('Embed', prog_id, 'rIdA')}{_object('Embed', prog_id, 'rIdB')}</w:p>"
        f"<w:p><w:r><w:t>Second</w:t></w:r>{_object('Embed', prog_id, 'rIdC')}</w:p>"
    )


def packages():
    """docx4j's ``packages()``: two OLE 1.0 packages in one paragraph, and a third in the next."""
    return (
        MiniDocx.body(_body("Package"))
        .part(
            "embeddings/oleObject1.bin",
            OLE_OBJECT,
            REL + "oleObject",
            "rIdA",
            ole10("notes.txt", "C:\\docs\\notes.txt", "C:\\tmp\\notes.txt", "NOTES_TEXT\r\n"),
        )
        .part(
            "embeddings/oleObject2.bin",
            OLE_OBJECT,
            REL + "oleObject",
            "rIdB",
            ole10("other.txt", "C:\\docs\\other.txt", "C:\\tmp\\other.txt", "OTHER_TEXT\r\n"),
        )
        .part(
            "embeddings/oleObject3.bin",
            OLE_OBJECT,
            REL + "oleObject",
            "rIdC",
            ole10("third.txt", "C:\\docs\\third.txt", "C:\\tmp\\third.txt", "THIRD_TEXT\r\n"),
        )
        .load()
    )


def _document(text: str) -> bytes:
    return MiniDocx.body(f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>").save()


def documents():
    """``packages()`` over embedded documents, which this package reads (section 5's departure)."""
    made = MiniDocx.body(_body("Word.Document.12"))
    for rel_id, name, text in (
        ("rIdA", "notes", "NOTES_TEXT"),
        ("rIdB", "other", "OTHER_TEXT"),
        ("rIdC", "third", "THIRD_TEXT"),
    ):
        made.part(f"embeddings/{name}.docx", DOCX, REL + "package", rel_id, _document(text))
    return made.load()


# --- the object (REQ-512) ------------------------------------------------------------------------------


def test_probe_objects_described(probe):
    """REQ-512's members on the probe's objects: a workbook, a document, an OLE 1.0 package, a control, a presentation."""
    runs = objects(extract(probe))
    assert len(runs) == 5
    xlsx, docx, ole, control, pptx = (r.object for r in runs)

    assert xlsx.kind == "embed"
    assert xlsx.prog_id == "Excel.Sheet.12"
    assert xlsx.part_name == "/word/embeddings/probe39.xlsx"
    assert xlsx.content_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert isinstance(xlsx.part, EmbeddedPackagePart)
    assert xlsx.part is probe.main_document_part.relationships_part.get_part(xlsx.relationship_id)
    assert xlsx.label is None
    assert type(xlsx.source).__name__ == "CTObject"

    assert docx.kind == "embed"
    assert docx.part_name == "/word/embeddings/probe40.docx"
    assert docx.prog_id == "Word.Document.12"

    # an OLE 1.0 package: its native data's names, as Tika hands such a package on
    assert ole.kind == "embed"
    assert ole.prog_id == "Package"
    assert ole.part_name == "/word/embeddings/probe41.bin"
    assert ole.content_type == OLE_OBJECT
    assert (ole.label, ole.file_name, ole.command) == ("probe.txt", "probe.txt", "probe.txt")

    # an ActiveX control (w:control in w:object): its part, no ProgID
    assert control.kind == "control"
    assert control.prog_id is None
    assert runs[3].prog_id is None
    assert control.part_name == "/word/activeX/activeX42.xml"
    assert control.content_type == ACTIVEX
    assert type(control.source).__name__ == "CTObject"

    assert pptx.prog_id == "PowerPoint.Show.12"
    assert pptx.part_name.endswith(".pptx")
    assert all(r.prog_id == r.object.prog_id for r in runs)


def test_package_names():
    """REQ-512: an OLE 1.0 package's label, file name and command, each where the native data has it."""
    described = objects(extract(packages()))[0].object
    assert described.label == "notes.txt"
    assert described.file_name == "C:\\docs\\notes.txt"
    assert described.command == "C:\\tmp\\notes.txt"
    assert described.relationship_id == "rIdA"
    assert isinstance(described.part, BinaryPart)
    # a Package whose binary holds no Ole10Native stream: no names, and no failure
    other = (
        MiniDocx.body(f"<w:p>{_object('Embed', 'Package', 'rIdX')}</w:p>")
        .part("embeddings/oleObject1.bin", OLE_OBJECT, REL + "oleObject", "rIdX", b"not an OLE file")
        .load()
    )
    none = objects(extract(other))[0].object
    assert (none.label, none.file_name, none.command) == (None, None, None)
    assert none.part_name == "/word/embeddings/oleObject1.bin"


def test_linked_object():
    """REQ-512: a linked object has its relationship id and ProgID, and no part; it is never read, nor asked about."""
    package = (
        MiniDocx.body(f"<w:p><w:r><w:t>Linked</w:t></w:r>{_object('Link', 'Excel.Sheet.12', 'rIdL')}</w:p>")
        .external("rIdL", REL + "oleObject", "file:///C:/data/book.xlsx")
        .load()
    )
    asked = []
    result = extract(package, embedded=True, embedded_filter=lambda o, a: asked.append(o))
    described = objects(result)[0].object
    assert described.kind == "link"
    assert described.relationship_id == "rIdL"
    assert described.prog_id == "Excel.Sheet.12"
    assert described.part is None
    assert described.part_name is None
    assert described.content_type is None
    assert asked == []
    assert plain_text(result) == "Linked\n"
    assert result.warnings == []


def test_a_link_whose_relationship_is_internal_has_no_part():
    """REQ-512: absent for a linked object, even one whose relationship is internal."""
    package = (
        MiniDocx.body(f"<w:p>{_object('Link', 'Package', 'rIdL')}</w:p>")
        .part("embeddings/oleObject1.bin", OLE_OBJECT, REL + "oleObject", "rIdL", ole10("a.txt", "a.txt", "a.txt", "A"))
        .load()
    )
    described = objects(extract(package, embedded=True))[0].object
    assert (described.kind, described.relationship_id, described.part, described.label) == ("link", "rIdL", None, None)


def test_a_relationship_that_does_not_resolve_has_no_part():
    package = MiniDocx.body(f"<w:p>{_object('Embed', 'Package', 'rIdNone')}</w:p>").load()
    result = extract(package, embedded=True)
    described = objects(result)[0].object
    assert (described.kind, described.relationship_id, described.part_name) == ("embed", "rIdNone", None)
    assert result.warnings == []


def test_a_type_that_is_absent_is_embed():
    """REQ-512: ``embed`` where ``o:OLEObject/@Type`` is absent."""
    xml = '<w:p><w:r><w:object><o:OLEObject ProgID="Package" r:id="rIdA"/></w:object></w:r></w:p>'
    package = MiniDocx.body(xml).load()
    assert objects(extract(package))[0].object.kind == "embed"


def test_controls():
    """REQ-512: a ``w:control`` standing in ``w:pict``, and an ``o:OLEObject`` whose part is an ActiveX control's, are controls."""
    ocx = (
        '<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" '
        'ax:classid="{8BD21D40-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistPropertyBag"/>'
    )
    package = (
        MiniDocx.body(
            '<w:p><w:r><w:pict><v:shape id="_x0000_i1025" style="width:20pt;height:10pt"/>'
            '<w:control r:id="rIdX" w:name="CheckBox1" w:shapeid="_x0000_i1025"/></w:pict></w:r>'
            f"{_object('Embed', 'Forms.CheckBox.1', 'rIdY')}</w:p>"
        )
        .part("activeX/activeX1.xml", ACTIVEX, REL + "control", "rIdX", ocx)
        .part("activeX/activeX2.xml", ACTIVEX, REL + "control", "rIdY", ocx)
        .load()
    )
    runs = objects(extract(package, embedded=True))
    assert len(runs) == 2
    pict, ole = runs[0].object, runs[1].object
    assert pict.kind == "control"
    assert pict.relationship_id == "rIdX"
    assert pict.part_name == "/word/activeX/activeX1.xml"
    assert pict.content_type == ACTIVEX
    assert type(pict.source).__name__ == "CTControl"
    assert ole.kind == "control"
    assert ole.prog_id == "Forms.CheckBox.1"
    assert ole.part_name == "/word/activeX/activeX2.xml"


def test_a_binary_activex_part_is_a_control():
    """The second of REQ-512's two content types: ``application/vnd.ms-office.activeX``."""
    package = (
        MiniDocx.body(f"<w:p>{_object('Embed', 'Forms.CheckBox.1', 'rIdY')}</w:p>")
        .part("activeX/activeX1.bin", "application/vnd.ms-office.activeX", REL + "control", "rIdY", b"\0")
        .load()
    )
    assert objects(extract(package))[0].object.kind == "control"


def test_a_wildcard_object_is_described_as_a_typed_one():
    """A kept branch's ``w:object`` is no root element (CR-012 Phase C), so a wildcard: its object, its control."""
    branch = (
        '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:object>'
        '<o:OLEObject Type="Link" ProgID="Excel.Sheet.12" r:id="rIdL"/></w:object></mc:Choice></mc:AlternateContent></w:r>'
        '<w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:object><w:control r:id="rIdX"/></w:object></mc:Choice>'
        "</mc:AlternateContent></w:r></w:p>"
    )
    package = (
        MiniDocx.body(branch)
        .external("rIdL", REL + "oleObject", "file:///C:/data/book.xlsx")
        .part("activeX/activeX1.xml", ACTIVEX, REL + "control", "rIdX", "<ax:ocx xmlns:ax='urn:ax'/>")
        .load()
    )
    link, control = (r.object for r in objects(extract(package)))
    assert (link.kind, link.relationship_id, link.prog_id, link.part) == ("link", "rIdL", "Excel.Sheet.12", None)
    assert (control.kind, control.relationship_id, control.prog_id) == ("control", "rIdX", None)
    assert control.part_name == "/word/activeX/activeX1.xml"
    assert type(link.source).__name__ == "AnyElement"


def test_only_an_object_run_carries_one():
    """Only an object run carries one: every other run's is None."""
    result = extract(MiniDocx.body("<w:p><w:r><w:t>Text</w:t></w:r></w:p>").load())
    assert result.blocks[0].runs[0].object is None


# --- the block model's JSON (section 4.3) --------------------------------------------------------------


def test_the_json_carries_the_object_member(probe):
    """REQ-512's members in this JSON's spelling, each left out where None; the part itself and the source are not in it."""
    runs = objects(extract(probe))
    assert runs[2].to_dict()["object"] == {
        "kind": "embed",
        "relationship_id": "rId30",
        "part": "/word/embeddings/probe41.bin",
        "content_type": OLE_OBJECT,
        "label": "probe.txt",
        "file_name": "probe.txt",
        "command": "probe.txt",
    }
    assert runs[2].to_dict()["prog_id"] == "Package"
    assert runs[3].to_dict()["object"] == {
        "kind": "control",
        "relationship_id": "rId31",
        "part": "/word/activeX/activeX42.xml",
        "content_type": ACTIVEX,
    }
    assert "prog_id" not in runs[3].to_dict()
    json.dumps([b.to_dict() for b in extract(probe).blocks])


def test_the_json_of_a_link_and_of_another_run():
    package = (
        MiniDocx.body(f"<w:p><w:r><w:t>Linked</w:t></w:r>{_object('Link', 'Excel.Sheet.12', 'rIdL')}</w:p>")
        .external("rIdL", REL + "oleObject", "file:///C:/data/book.xlsx")
        .load()
    )
    text, link = extract(package).blocks[0].runs
    assert link.to_dict()["object"] == {"kind": "link", "relationship_id": "rIdL"}
    assert "object" not in text.to_dict()


def test_str():
    """docx4j's ``toString``."""
    assert str(TextObject("link", prog_id="Excel.Sheet.12")) == "link Excel.Sheet.12"
    assert str(objects(extract(packages()))[0].object) == "embed Package /word/embeddings/oleObject1.bin (notes.txt)"
    assert str(TextObject("control")) == "control"


# --- the filter (REQ-522) ------------------------------------------------------------------------------


def test_no_filter_reads_every():
    """REQ-522: no filter reads every object (REQ-512), as a None one does."""
    every = "First\nNOTES_TEXT\nOTHER_TEXT\nSecond\nTHIRD_TEXT\n"
    assert plain_text(extract(documents(), embedded=True)) == every
    assert plain_text(extract(documents(), TextOptions(embedded=True, embedded_filter=None))) == every
    assert TextOptions().embedded_filter is None


def test_filter_declines_one_of_two():
    """REQ-522: asked once per object, before the reading; a declined object contributes nothing and keeps its run."""
    package = documents()
    asked: dict[str, TextObject] = {}

    def reads(described: TextObject, address: str) -> bool:
        asked[address] = described
        return described.part_name != "/word/embeddings/notes.docx"

    result = extract(package, embedded=True, embedded_filter=reads)
    assert list(asked) == ["body/0/object:0", "body/0/object:1", "body/1/object:0"]
    assert plain_text(result) == "First\nOTHER_TEXT\nSecond\nTHIRD_TEXT\n"
    # the declined object keeps its run and its description, the one the filter was given
    runs = objects(result)
    assert len(runs) == 3
    assert runs[0].object is asked["body/0/object:0"]
    assert runs[1].object is asked["body/0/object:1"]
    # the second object's content is where it would be had the first been read
    assert [b.address for b in result.blocks if b.story == "EMBEDDED"] == [
        "body/0/object:1/body/0",
        "body/1/object:0/body/0",
    ]
    assert result.warnings == []
    # without embedded: true nothing is read, and the filter is never asked
    asked.clear()
    extract(package, embedded_filter=reads)
    assert not asked


def test_filter_over_the_probe(probe):
    """REQ-522: a host that reads workbooks itself; one level, so the embedded document's own object is never asked about."""
    asked: list[str | None] = []

    def reads(described: TextObject, address: str) -> bool:
        asked.append(described.prog_id)
        return "spreadsheetml" not in described.content_type

    result = extract(probe, embedded=True, embedded_filter=reads)
    assert asked == ["Excel.Sheet.12", "Word.Document.12", "Package", "PowerPoint.Show.12"]
    text = plain_text(result)
    assert "EMB39_XLSX_CELL" not in text
    assert "EMB40_DOCX_TEXT" in text
    assert "EMB40_NESTED_TEXT" not in text
    assert objects(result)[0].prog_id == "Excel.Sheet.12"
    # the departure: docx4j reads these two, which this package defers (CR-011 Q7); admitted, they are content.not_read
    not_read = [w.part_name for w in result.warnings if w.kind == TextWarning.CONTENT_NOT_READ]
    assert not_read == ["/word/embeddings/probe41.bin", "/word/embeddings/probe53.pptx"]


def test_failing_filter_declines():
    """REQ-522: a filter that fails declines, with ``content.not_read`` naming the object's part at the paragraph."""

    def reads(described: TextObject, address: str) -> bool:
        if address == "body/0/object:1":
            raise RuntimeError("host failure")
        return True

    result = extract(documents(), embedded=True, embedded_filter=reads)
    assert plain_text(result) == "First\nNOTES_TEXT\nSecond\nTHIRD_TEXT\n"
    assert [(w.kind, w.address, w.part_name) for w in result.warnings] == [
        (TextWarning.CONTENT_NOT_READ, "body/0", "/word/embeddings/other.docx")
    ]
    assert result.warnings[0].selector == "/word/embeddings/other.docx"
    assert len(objects(result)) == 3


def test_the_filter_over_ole_packages_is_asked_as_docx4j_asks_it():
    """docx4j's ``filterDeclinesOneOfTwo`` over its packages: its asks and declines; each admitted is deferred (Q4)."""
    asked: dict[str, str | None] = {}

    def reads(described: TextObject, address: str) -> bool:
        asked[address] = described.label
        return described.label != "notes.txt"

    result = extract(packages(), embedded=True, embedded_filter=reads)
    assert asked == {
        "body/0/object:0": "notes.txt",
        "body/0/object:1": "other.txt",
        "body/1/object:0": "third.txt",
    }
    assert plain_text(result) == "First\nSecond\n"
    assert [(w.kind, w.address, w.part_name) for w in result.warnings] == [
        (TextWarning.CONTENT_NOT_READ, "body/0", "/word/embeddings/oleObject2.bin"),
        (TextWarning.CONTENT_NOT_READ, "body/1", "/word/embeddings/oleObject3.bin"),
    ]


def test_a_declined_ole_package_raises_nothing():
    """So a host that takes OLE 1.0 packages itself sees no warning for them."""
    result = extract(packages(), embedded=True, embedded_filter=lambda o, a: False)
    assert plain_text(result) == "First\nSecond\n"
    assert result.warnings == []
    assert len(objects(result)) == 3


def test_an_object_that_would_not_be_read_is_not_asked_about():
    """Only an embedded package and an OLE 1.0 package: a plain OLE binary, and a control, take no index."""
    package = (
        MiniDocx.body(
            f"<w:p>{_object('Embed', 'Equation.3', 'rIdQ')}{_object('Embed', 'Forms.CheckBox.1', 'rIdX')}"
            f"{_object('Embed', 'Word.Document.12', 'rIdD')}</w:p>"
        )
        .part("embeddings/oleObject1.bin", OLE_OBJECT, REL + "oleObject", "rIdQ", b"\0")
        .part("activeX/activeX1.xml", ACTIVEX, REL + "control", "rIdX", "<ax:ocx xmlns:ax='urn:ax'/>")
        .part("embeddings/doc.docx", DOCX, REL + "package", "rIdD", _document("DOC_TEXT"))
        .load()
    )
    asked: list[str] = []
    result = extract(package, embedded=True, embedded_filter=lambda o, a: asked.append(a) or True)
    assert asked == ["body/0/object:0"]
    assert [r.object.kind for r in objects(result)] == ["embed", "control", "embed"]


def test_the_filter_is_not_asked_one_level_down():
    """An embedded document's own objects are never read, so never offered."""
    inner = MiniDocx.body(f"<w:p>{_object('Embed', 'Word.Document.12', 'rIdI')}</w:p>")
    inner.part("embeddings/inner.docx", DOCX, REL + "package", "rIdI", _document("INNER_TEXT"))
    package = (
        MiniDocx.body(f"<w:p>{_object('Embed', 'Word.Document.12', 'rIdO')}</w:p>")
        .part("embeddings/outer.docx", DOCX, REL + "package", "rIdO", inner.save())
        .load()
    )
    asked: list[str] = []
    result = extract(package, embedded=True, embedded_filter=lambda o, a: asked.append(a) or True)
    assert asked == ["body/0/object:0"]
    assert "INNER_TEXT" not in plain_text(result)
    # the object one level down still has its run and its description
    inner_runs = objects(result)
    assert [r.object.part_name for r in inner_runs] == ["/word/embeddings/outer.docx", "/word/embeddings/inner.docx"]


# --- the option ----------------------------------------------------------------------------------------


def test_the_option_is_set_by_name_and_is_no_section_10_name():
    def admits(described: TextObject, address: str) -> bool:
        return True

    options = TextOptions().set("embedded_filter", admits)
    assert options.embedded_filter is admits
    assert options.copy().embedded_filter is admits
    assert TextOptions.of({"embedded_filter": admits}).embedded_filter is admits
    assert "embedded_filter" not in options.to_dict()
    assert options.set("embedded_filter", None).embedded_filter is None


def test_a_filter_that_is_not_callable_is_refused():
    with pytest.raises(TextOptionError) as caught:
        TextOptions(embedded_filter="no")
    assert (caught.value.option, caught.value.statement, caught.value.reason) == (
        "embedded_filter",
        "REQ-522",
        "invalid",
    )


# --- altchunk.not_expanded's part name (REQ-513, REQ-801) ----------------------------------------------


def test_chunk_warning_names_the_part(probe):
    """The warning carries the chunk's content type, and its part name as its selector; a nested chunk names the chunk holding it."""
    by_part = {w.part_name: w for w in extract(probe).warnings if w.kind == TextWarning.ALTCHUNK_NOT_EXPANDED}
    rtf = by_part["/word/chunk45.rtf"]
    assert rtf.content_type == "application/rtf"
    assert rtf.selector == "/word/chunk45.rtf"
    nested = by_part["/word/chunk47.docx"]
    assert nested.content_type == "text/html"
    assert nested.address.split("/")[0] == "body" and "/altchunk:0/body/" in nested.address
    # altchunks: false: every chunk, by its own part
    for warning in extract(probe, altchunks=False).warnings:
        if warning.kind == TextWarning.ALTCHUNK_NOT_EXPANDED:
            assert warning.selector.startswith("/word/")
            assert warning.selector == warning.part_name


# --- a reading changes nothing (REQ-204) ---------------------------------------------------------------


def test_a_reading_leaves_every_part_as_it_was():
    package = packages()
    before = {name: bytes(part.data) for name, part in _binaries(package)}
    extract(package, embedded=True, embedded_filter=lambda o, a: True)
    assert {name: bytes(part.data) for name, part in _binaries(package)} == before
    assert not package.main_document_part.is_unmarshalled


def _binaries(package):
    return [(str(p.part_name), p) for p in package.parts.values() if isinstance(p, BinaryPart)]
