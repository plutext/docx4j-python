"""Parts, relationships, ``create_package`` and editing a package.

CR-002 sections 5.2, 5.4 and 5.5.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_parts_and_relationships.py -q
"""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py.openpackaging import (  # noqa: E402
    PAGE_SIZES,
    AddPartBehaviour,
    ContentTypes,
    Docx4JException,
    ImagePart,
    InvalidOperationException,
    MainDocumentPart,
    Namespaces,
    OpcPackage,
    PartName,
    RelationshipsPart,
    StyleDefinitionsPart,
    WordprocessingMLPackage,
)

TOC = ROOT / "samples" / "toc.docx"
IMAGES = ROOT / "samples" / "Images.docx"


@pytest.fixture
def png() -> bytes:
    with zipfile.ZipFile(IMAGES) as zf:
        name = next(n for n in zf.namelist() if n.lower().endswith(".png"))
        return zf.read(name)


def entries(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


# ---------------------------------------------------------------------------
# the relationship graph as loaded
# ---------------------------------------------------------------------------


def test_the_package_relationships():
    with OpcPackage.load(TOC) as pkg:
        rels = pkg.relationships_part
        assert rels.is_package_relationship_part
        assert rels.source_name == PartName.ROOT
        assert rels.part_name.name == "/_rels/.rels"
        assert rels.get_relationship_by_type(Namespaces.DOCUMENT) is not None
        assert len(rels) >= 3


def test_get_part_by_relationship_and_back():
    with WordprocessingMLPackage.load(TOC) as pkg:
        main = pkg.main_document_part
        rels = main.relationships_part
        rel = rels.get_relationship_by_type(Namespaces.STYLES)
        styles = rels.get_part(rel)
        assert isinstance(styles, StyleDefinitionsPart)
        assert rels.get_rel(styles.part_name) is rel
        assert rels.is_a_target(styles.part_name)
        assert rels.get_part(rel.id) is styles


def test_external_relationships_are_not_parts():
    from docx4j_py.openpackaging import is_external

    with OpcPackage.load(ROOT / "samples" / "tables.docx") as pkg:
        rels = pkg.main_document_part.relationships_part
        external = [r for r in rels.list if is_external(r)]
        assert external, "the fixture must hold an external relationship"
        for rel in external:
            assert rels.get_part(rel) is None
            assert pkg.get_part(PartName.of("/word/media/x.png")) is None


def test_custom_xml_parts_are_indexed_by_item_id():
    with OpcPackage.load(TOC) as pkg:
        assert pkg.custom_xml_data_storage_parts
        for item_id, part in pkg.custom_xml_data_storage_parts.items():
            assert item_id == item_id.lower()
            assert part.part_name.name.startswith("/customXml/")


# ---------------------------------------------------------------------------
# ids
# ---------------------------------------------------------------------------


def test_the_next_free_id_skips_the_used_ones():
    with WordprocessingMLPackage.load(TOC) as pkg:
        rels = pkg.main_document_part.relationships_part
        used = {r.id for r in rels.list}
        rels.reset_id_allocator()
        for _ in range(5):
            assert rels.get_next_id() not in used


def test_a_duplicate_id_is_refused():
    from docx4j_py.relationships import Relationship

    with WordprocessingMLPackage.load(TOC) as pkg:
        rels = pkg.main_document_part.relationships_part
        taken = rels.list[0].id
        with pytest.raises(InvalidOperationException, match="Refusing"):
            rels.add_relationship(
                Relationship(id=taken, type_value=Namespaces.IMAGE, target="media/x.png")
            )


def test_an_external_relationship_round_trips():
    with WordprocessingMLPackage.load(TOC) as pkg:
        rels = pkg.main_document_part.relationships_part
        rel = rels.add_external_relationship(Namespaces.HYPERLINK, "https://example.com/")
        assert rel.id
        data = pkg.save()
    with WordprocessingMLPackage.load(data) as again:
        rels = again.main_document_part.relationships_part
        found = rels.get_relationship_by_id(rel.id)
        assert found is not None
        assert found.target == "https://example.com/"
        assert found.target_mode.value == "External"


# ---------------------------------------------------------------------------
# adding, renaming and removing a part
# ---------------------------------------------------------------------------


def test_add_an_image_part_with_a_relationship(png):
    with WordprocessingMLPackage.load(TOC) as pkg:
        main = pkg.main_document_part
        image = ImagePart("/word/media/image1.png")
        image.set_bytes(png)
        rel = main.add_target_part(image)

        assert rel.type_value == Namespaces.IMAGE
        assert rel.target == "media/image1.png"
        assert pkg.get_part("/word/media/image1.png") is image
        assert image.package is pkg
        assert (
            pkg.content_type_manager.get_content_type("/word/media/image1.png")
            == ContentTypes.IMAGE_PNG
        )
        data = pkg.save()

    saved = entries(data)
    assert saved["word/media/image1.png"] == png
    with WordprocessingMLPackage.load(data) as again:
        reloaded = again.get_part("/word/media/image1.png")
        assert isinstance(reloaded, ImagePart)
        assert reloaded.data == png


def test_reuse_existing_returns_the_part_already_there(png):
    with WordprocessingMLPackage.load(TOC) as pkg:
        main = pkg.main_document_part
        first = ImagePart("/word/media/image1.png")
        first.set_bytes(png)
        main.add_target_part(first)

        second = ImagePart("/word/media/image1.png")
        second.set_bytes(b"different")
        rel = main.add_target_part(second, AddPartBehaviour.REUSE_EXISTING)

        assert pkg.get_part("/word/media/image1.png") is first
        assert main.relationships_part.get_part(rel) is first


def test_rename_if_name_exists_allocates_a_new_name(png):
    with WordprocessingMLPackage.load(TOC) as pkg:
        main = pkg.main_document_part
        first = ImagePart("/word/media/image1.png")
        first.set_bytes(png)
        main.add_target_part(first)

        second = ImagePart("/word/media/image1.png")
        second.set_bytes(b"\x89PNG\r\n\x1a\n second")
        main.add_target_part(second, AddPartBehaviour.RENAME_IF_NAME_EXISTS)

        # docx4j's own answer: the prefix is everything before the extension,
        # so `image1.png` becomes `image12.png` and not `image2.png`. Ugly, and
        # kept, because a docx4j user's part names have to come out the same.
        assert second.part_name.name == "/word/media/image12.png"
        assert pkg.get_part("/word/media/image1.png") is first
        assert pkg.get_part("/word/media/image12.png") is second
        data = pkg.save()
    assert "word/media/image12.png" in entries(data)


def test_renaming_a_part_re_keys_it_and_moves_its_bytes(png):
    with WordprocessingMLPackage.load(TOC) as pkg:
        main = pkg.main_document_part
        image = ImagePart("/word/media/image1.png")
        image.set_bytes(png)
        rel = main.add_target_part(image)

        image.part_name = "/word/media/logo.png"
        rel.target = "media/logo.png"
        pkg.content_type_manager.add_content_type(image.part_name, image.content_type)
        assert pkg.get_part("/word/media/logo.png") is image
        assert pkg.get_part("/word/media/image1.png") is None
        data = pkg.save()
    saved = entries(data)
    assert "word/media/logo.png" in saved
    assert "word/media/image1.png" not in saved


def test_remove_a_part_and_its_relationship(png):
    with WordprocessingMLPackage.load(TOC) as pkg:
        main = pkg.main_document_part
        image = ImagePart("/word/media/image1.png")
        image.set_bytes(png)
        main.add_target_part(image)

        removed = main.relationships_part.remove_part("/word/media/image1.png")
        assert [n.name for n in removed] == ["/word/media/image1.png"]
        assert pkg.get_part("/word/media/image1.png") is None
        assert main.relationships_part.get_rel("/word/media/image1.png") is None
        data = pkg.save()
    assert "word/media/image1.png" not in entries(data)


def test_removing_the_styles_part_from_a_real_document():
    with WordprocessingMLPackage.load(TOC) as pkg:
        main = pkg.main_document_part
        styles = pkg.style_definitions_part
        main.relationships_part.remove_part(styles.part_name)
        data = pkg.save()
    assert "word/styles.xml" not in entries(data)
    with WordprocessingMLPackage.load(data) as again:
        assert again.style_definitions_part is None


# ---------------------------------------------------------------------------
# create_package
# ---------------------------------------------------------------------------


def test_create_package_makes_what_docx4j_makes():
    pkg = WordprocessingMLPackage.create_package()
    names = sorted(n.name for n in pkg.parts)
    assert names == [
        "/docProps/app.xml",
        "/docProps/core.xml",
        "/word/document.xml",
        "/word/settings.xml",
        "/word/styles.xml",
    ]
    assert isinstance(pkg.main_document_part, MainDocumentPart)
    # styles hangs off the document part, not off the package, as in docx4j
    assert pkg.relationships_part.get_rel("/word/styles.xml") is None
    assert pkg.main_document_part.relationships_part.get_rel("/word/styles.xml")


@pytest.mark.parametrize("page_size", sorted(PAGE_SIZES))
@pytest.mark.parametrize("landscape", [False, True])
def test_create_package_page_sizes(page_size, landscape):
    pkg = WordprocessingMLPackage.create_package(page_size=page_size, landscape=landscape)
    sect_pr = pkg.main_document_part.contents.body.sect_pr
    width, height, code = PAGE_SIZES[page_size]
    assert sect_pr.pg_sz.code == code
    if landscape:
        assert (sect_pr.pg_sz.w, sect_pr.pg_sz.h) == (height, width)
        assert sect_pr.pg_sz.orient == "landscape"
    else:
        assert (sect_pr.pg_sz.w, sect_pr.pg_sz.h) == (width, height)
        assert sect_pr.pg_sz.orient is None
    assert (sect_pr.pg_mar.top, sect_pr.pg_mar.left) == (1440, 1440)


def test_create_package_rejects_an_unknown_page_size():
    with pytest.raises(ValueError, match="Unknown page size"):
        WordprocessingMLPackage.create_package(page_size="A0")


def test_a_created_document_saves_and_reloads():
    from docx4j_py.wml import p, text_of

    pkg = WordprocessingMLPackage.create_package()
    pkg.main_document_part.contents.body.content.append(p("Hello World"))
    data = pkg.save()

    with WordprocessingMLPackage.load(data) as again:
        assert text_of(again.main_document_part.contents) == "Hello World"
        assert again.style_definitions_part is not None
        assert len(again.style_definitions_part.contents.style) > 10
        settings = again.document_settings_part.contents
        names = [c.name for c in settings.compat.compat_setting]
        assert "overrideTableStyleFontSizeAndJustification" in names
        assert again.skipped == []


def test_the_default_styles_resource_is_docx4js():
    from docx4j_py.openpackaging.resources import DEFAULT_PARTS, default_part_bytes

    for name in DEFAULT_PARTS:
        data = default_part_bytes(name)
        assert data.startswith(b"<?xml")
        assert b'encoding="UTF-8"' in data[:80] or b"encoding='UTF-8'" in data[:80]
    with pytest.raises(KeyError):
        default_part_bytes("nope.xml")


def test_the_default_numbering_and_fonts_unmarshal():
    from docx4j_py.openpackaging import FontTablePart, NumberingDefinitionsPart

    numbering = NumberingDefinitionsPart()
    assert numbering.unmarshal_default_numbering().abstract_num
    assert numbering.skipped == []

    fonts = FontTablePart()
    assert fonts.unmarshal_default_fonts().font
    assert fonts.skipped == []


# ---------------------------------------------------------------------------
# parts, the mapping
# ---------------------------------------------------------------------------


def test_parts_is_case_insensitive():
    with OpcPackage.load(TOC) as pkg:
        assert pkg.get_part("/WORD/DOCUMENT.XML") is pkg.get_part("/word/document.xml")
        assert "/word/document.xml" in pkg.parts
        assert pkg.get_part("/word/nothing.xml") is None
        assert len(pkg.parts) == len(pkg.parts.parts())


def test_parts_of_type():
    from docx4j_py.openpackaging import XmlPart

    with OpcPackage.load(TOC) as pkg:
        assert len(pkg.parts.of_type(XmlPart)) > 5


# ---------------------------------------------------------------------------
# mc:Ignorable on write, CR-002 section 5.6
# ---------------------------------------------------------------------------

#: The six parts CR-001 section 13 identified as declaring ``mc:Ignorable``
#: prefixes the fixed prefix table does not know (`w16se`, `w16cid`).
IGNORABLE_PARTS = [
    ("2016_image_with_text_effects.docx", "/word/document.xml"),
    ("2016_image_with_text_effects.docx", "/word/styles.xml"),
    ("2016_image_with_text_effects.docx", "/word/fontTable.xml"),
    ("DrawingML_GraphicData_wps.docx", "/word/document.xml"),
    ("DrawingML_GraphicData_wps.docx", "/word/styles.xml"),
    ("DrawingML_GraphicData_wps.docx", "/word/fontTable.xml"),
]


@pytest.mark.parametrize(("sample", "part_name"), IGNORABLE_PARTS)
def test_every_ignorable_prefix_is_declared_on_the_re_serialised_root(sample, part_name):
    from lxml import etree

    with OpcPackage.load(ROOT / "samples" / sample) as pkg:
        part = pkg.get_part(part_name)
        part.contents  # noqa: B018
        root = etree.fromstring(part.xml)

    ignorable = root.get(f"{{{Namespaces.MARKUP_COMPATIBILITY}}}Ignorable")
    assert ignorable, f"{sample} {part_name} must carry mc:Ignorable"
    declared = set(root.nsmap)
    missing = [p for p in ignorable.split() if p not in declared]
    assert missing == [], f"{sample} {part_name}: {missing} named but not declared"


def test_a_prefix_the_table_does_not_know_is_remembered_from_the_source():
    """``w16se`` and ``w16cid`` are declared by Word and not by the table."""
    from lxml import etree

    with OpcPackage.load(ROOT / "samples" / "2016_image_with_text_effects.docx") as pkg:
        part = pkg.get_part("/word/document.xml")
        assert part.source_ns_map == {}, "nothing is read before the unmarshal"
        part.contents  # noqa: B018
        assert "w16se" in part.source_ns_map or "w16cid" in part.source_ns_map
        root = etree.fromstring(part.xml)
    for prefix in ("w16se", "w16cid"):
        if prefix in part.source_ns_map:
            assert root.nsmap.get(prefix) == part.source_ns_map[prefix]


def test_the_relationships_namespace_is_the_default_one_on_write():
    """Word writes ``<Relationships xmlns="...">``, not ``<rel:Relationships>``."""
    from lxml import etree

    with OpcPackage.load(TOC) as pkg:
        data = pkg.relationships_part.xml
    root = etree.fromstring(data)
    assert root.nsmap == {None: Namespaces.RELATIONSHIPS}
    assert root.tag == f"{{{Namespaces.RELATIONSHIPS}}}Relationships"


# ---------------------------------------------------------------------------
# errors
# ---------------------------------------------------------------------------


def test_a_zip_without_content_types_is_refused(tmp_path):
    from docx4j_py.openpackaging import InvalidFormatException

    broken = tmp_path / "broken.docx"
    with zipfile.ZipFile(broken, "w") as zf:
        zf.writestr("word/document.xml", "<x/>")
    with pytest.raises(InvalidFormatException, match=r"\[Content_Types\].xml"):
        OpcPackage.load(broken)


def test_something_that_is_not_a_zip_is_refused():
    with pytest.raises(Docx4JException, match="Not a zip"):
        OpcPackage.load(b"not a zip at all")


def test_load_refuses_a_type_it_cannot_make_a_store_of():
    with pytest.raises(Docx4JException, match="Cannot load a package"):
        OpcPackage.load(42)


def test_a_pptx_is_not_a_word_package():
    from docx4j_py.openpackaging import InvalidFormatException

    with pytest.raises(InvalidFormatException, match="Not a WordprocessingML"):
        WordprocessingMLPackage.load(ROOT / "samples" / "loadAndSave.pptx")


def test_a_relationships_part_has_no_relationships_part_of_its_own():
    with OpcPackage.load(TOC) as pkg:
        rels = pkg.relationships_part
        assert isinstance(rels, RelationshipsPart)
        assert rels.get_relationships_part(True) is None
