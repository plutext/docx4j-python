"""``insert_ooxml`` and ``FlatOpcStore``. CR-003 Phase C, sections 3.2 and 4.

The fixture is built rather than committed: :func:`flat_opc` turns a saved
package into the ``pkg:package`` document Word's clipboard produces, so the
test's incoming package is made with this engine's own verbs and carries an
image part, a style reference and a table.
"""

from __future__ import annotations

import base64
import io
import zipfile

import pytest
from conftest import SAMPLES, part_bytes, reloaded, sample

from docx4j_py import create_package, load
from docx4j_py.model.content import ContentError, Paragraph, Table
from docx4j_py.openpackaging.stores import FlatOpcStore

PKG = FlatOpcStore.NAMESPACE


def flat_opc(data: bytes) -> str:
    """A saved package as a flat OPC ``pkg:package`` document, as Word writes one."""
    from lxml import etree

    root = etree.Element(f"{{{PKG}}}package", nsmap={"pkg": PKG})
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        types = etree.fromstring(archive.read("[Content_Types].xml"))
        defaults = {
            element.get("Extension").lower(): element.get("ContentType")
            for element in types
            if element.tag.endswith("Default")
        }
        overrides = {
            element.get("PartName"): element.get("ContentType")
            for element in types
            if element.tag.endswith("Override")
        }
        for name in archive.namelist():
            if name == "[Content_Types].xml":
                continue
            part_name = "/" + name
            content_type = overrides.get(part_name) or defaults.get(
                name.rpartition(".")[2].lower(), "application/octet-stream"
            )
            part = etree.SubElement(root, f"{{{PKG}}}part")
            part.set(f"{{{PKG}}}name", part_name)
            part.set(f"{{{PKG}}}contentType", content_type)
            payload = archive.read(name)
            if "xml" in content_type:
                etree.SubElement(part, f"{{{PKG}}}xmlData").append(etree.fromstring(payload))
            else:
                binary = etree.SubElement(part, f"{{{PKG}}}binaryData")
                binary.text = base64.b64encode(payload).decode("ascii")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True).decode()


@pytest.fixture(scope="module")
def png() -> bytes:
    with zipfile.ZipFile(SAMPLES / "Images.docx") as archive:
        name = next(n for n in archive.namelist() if n.lower().endswith(".png"))
        return archive.read(name)


@pytest.fixture(scope="module")
def package_xml(png) -> str:
    """A ``pkg:package`` with a heading, a picture and a table."""
    source = create_package()
    source.id_seed = 5
    source.body.insert_paragraph("Pasted heading", style="Heading 1")
    source.body.insert_inline_picture(png, width=40, alt_text_description="pasted")
    source.body.insert_table(2, 2, values=[["a", "b"], ["c", "d"]])
    return flat_opc(source.save())


# ---------------------------------------------------------------------------
# the store
# ---------------------------------------------------------------------------


def test_the_flat_opc_store_is_a_part_store(package_xml):
    from docx4j_py.openpackaging.stores import PartStore

    store = FlatOpcStore.parse(package_xml)
    assert isinstance(store, PartStore)
    assert "word/document.xml" in store.part_names()
    assert "word/media/image1.png" in store.part_names()
    # a flat package carries no [Content_Types].xml; one is synthesised
    assert store.has("[Content_Types].xml")
    assert b"/word/media/image1.png" in store.load("[Content_Types].xml")
    assert store.size("word/document.xml") == len(store.load("word/document.xml"))
    assert repr(store).startswith("FlatOpcStore(")

    package = load(store)
    assert package.body.text.startswith("Pasted heading")


def test_the_store_refuses_what_is_not_a_package():
    from docx4j_py.openpackaging.exceptions import Docx4JException

    with pytest.raises(Docx4JException) as raised:
        FlatOpcStore.parse("<w:p xmlns:w='x'/>")
    assert "Not a flat OPC package" in str(raised.value)

    with pytest.raises(Docx4JException):
        FlatOpcStore.parse("<not well formed")


# ---------------------------------------------------------------------------
# inserting
# ---------------------------------------------------------------------------


def test_insert_ooxml_copies_the_parts_and_rewrites_the_references(new_package, package_xml, png):
    body = new_package.body
    body.insert_paragraph("Before")
    body.insert_inline_picture(png, width=10)  # image1.png is taken

    views = body.insert_ooxml(package_xml)

    assert [type(view).__name__ for view in views] == ["Paragraph", "Paragraph", "Table"]
    assert isinstance(views[0], Paragraph)
    assert views[0].text == "Pasted heading"
    assert views[0].style_id == "Heading1"
    assert isinstance(views[2], Table)
    assert views[2].values == [["a", "b"], ["c", "d"]]

    # the incoming image went in under a free name with a fresh relationship id
    pasted = body.inline_pictures[-1]
    assert str(pasted.image_part.part_name) == "/word/media/image2.png"
    assert pasted.get_bytes() == png
    assert pasted.alt_text_description == "pasted"
    assert pasted.rel_id != body.inline_pictures[0].rel_id
    assert new_package.main_document_part.relationships_part.get_part(pasted.rel_id) is (
        pasted.image_part
    )

    # and the report names what it wrote
    change = new_package.last_change
    assert change.operation == "insert_ooxml"
    assert "/word/media/image2.png" in change.parts_touched

    back = reloaded(new_package)
    assert back.body.text.startswith("Before")
    assert back.body.tables[0].values == [["a", "b"], ["c", "d"]]
    assert len(back.body.inline_pictures) == 2
    assert back.body.inline_pictures[1].get_bytes() == png


def test_numeric_ids_are_never_rewritten(new_package, package_xml):
    new_package.body.insert_ooxml(package_xml)
    picture = new_package.body.inline_pictures[0]

    # wp:docPr/@id is a number, not a relationship id, and keeps its value
    assert isinstance(picture.inline.doc_pr.id, int)
    assert picture.inline.doc_pr.id == 1


def test_a_bare_fragment_is_accepted_too(new_package):
    views = new_package.body.insert_ooxml(
        "<w:p><w:r><w:t>plain fragment</w:t></w:r></w:p>"
        "<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"
    )
    assert [type(view).__name__ for view in views] == ["Paragraph", "Table"]
    assert new_package.body.text.startswith("plain fragment")


def test_styles_and_numbering_are_not_merged(new_package, package_xml):
    from docx4j_py.model.markdown.importer import style_ids_of

    before = style_ids_of(new_package.style_definitions_part)
    new_package.body.insert_ooxml(package_xml)
    after = style_ids_of(new_package.style_definitions_part)

    assert after == before, "CR-003 section 4: that is docx4j's MergeDocx, not this"


def test_insert_ooxml_at_a_location_and_replace(new_package, package_xml):
    body = new_package.body
    body.insert_paragraph("only")

    body.insert_ooxml("<w:p><w:r><w:t>first</w:t></w:r></w:p>", location="Start")
    assert [p.text for p in body.paragraphs] == ["first", "only"]

    body.insert_ooxml(
        "<w:p><w:r><w:t>between</w:t></w:r></w:p>", location="After", target=body[0]
    )
    assert [p.text for p in body.paragraphs] == ["first", "between", "only"]

    body.insert_ooxml("<w:p><w:r><w:t>all there is</w:t></w:r></w:p>", location="Replace")
    assert body.text == "all there is"


def test_paragraph_level_merges_a_single_paragraph(new_package):
    paragraph = new_package.body.insert_paragraph("start")

    merged = paragraph.insert_ooxml("<w:p><w:r><w:t> and end</w:t></w:r></w:p>", location="End")
    assert merged == [paragraph]
    assert paragraph.text == "start and end"
    assert len(new_package.body) == 1

    beside = paragraph.insert_ooxml("<w:p><w:r><w:t>after</w:t></w:r></w:p>")
    assert beside[0].text == "after"
    assert len(new_package.body) == 2


def test_range_level_inserts_around_the_span(new_package):
    paragraph = new_package.body.insert_paragraph("the quick brown fox")
    span = paragraph.search("quick")[0]

    span.insert_ooxml("<w:p><w:r><w:t>slow</w:t></w:r></w:p>")
    assert paragraph.text == "the slow brown fox"

    other = paragraph.search("brown")[0]
    other.insert_ooxml("<w:p><w:r><w:t>very </w:t></w:r></w:p>", location="Before")
    assert paragraph.text == "the slow very brown fox"

    assert reloaded(new_package).body.text == "the slow very brown fox"


def test_a_package_with_no_part_to_relate_to_is_refused(package_xml):
    from docx4j_py.model.content import Body
    from docx4j_py.wml import Body as BodyElement

    body = Body(None, BodyElement())
    with pytest.raises(ContentError) as raised:
        body.insert_ooxml(package_xml)
    assert raised.value.code == "ooxml.no_part"


def test_a_package_with_no_main_part_says_so(new_package):
    from lxml import etree

    root = etree.Element(f"{{{PKG}}}package", nsmap={"pkg": PKG})
    part = etree.SubElement(root, f"{{{PKG}}}part")
    part.set(f"{{{PKG}}}name", "/_rels/.rels")
    part.set(
        f"{{{PKG}}}contentType",
        "application/vnd.openxmlformats-package.relationships+xml",
    )
    rels = "http://schemas.openxmlformats.org/package/2006/relationships"
    etree.SubElement(part, f"{{{PKG}}}xmlData").append(etree.Element(f"{{{rels}}}Relationships"))
    xml = etree.tostring(root, encoding="unicode")

    with pytest.raises(ContentError) as raised:
        new_package.body.insert_ooxml(xml)
    assert raised.value.code == "ooxml.no_main_part"


def test_only_the_touched_parts_are_rebuilt(package_xml):
    package = sample("2010-sample1.docx")
    _ = package.body
    before = part_bytes(package.save())

    package.body.insert_ooxml(package_xml)
    after = part_bytes(package.save())

    assert "word/media/image1.png" in after
    for name, data in before.items():
        if name not in ("word/document.xml", "[Content_Types].xml"):
            assert after[name] == data, name


def test_a_dry_run_of_insert_ooxml_un_adds_what_it_copied(package_xml):
    package = sample("2010-sample1.docx")
    _ = package.body
    before = part_bytes(package.save(), relationships=True)

    with package.dry_run() as trial:
        trial.body.insert_ooxml(package_xml)
        assert package.get_part("/word/media/image1.png") is not None

    assert package.get_part("/word/media/image1.png") is None
    assert part_bytes(package.save(), relationships=True) == before
