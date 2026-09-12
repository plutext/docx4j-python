"""``[Content_Types].xml`` and the part registry. CR-002 sections 5.3 and 5.4.

.venv-fork/bin/python -m pytest tests/openpackaging/test_content_types.py -q
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py.openpackaging import (  # noqa: E402
    ContentTypeManager,
    ContentTypes,
    ImagePart,
    InvalidFormatException,
    MainDocumentPart,
    Namespaces,
    PartName,
    StyleDefinitionsPart,
    ThemePart,
    default_part_registry,
    is_stored_uncompressed,
    is_xml_content_type,
)
from docx4j_py.openpackaging.content_types import CONTENT_TYPES_NS  # noqa: E402
from docx4j_py.openpackaging.parts.binary_part import (  # noqa: E402
    AlternativeFormatInputPart,
    BinaryPart,
    EmbeddedPackagePart,
    ObfuscatedFontPart,
    OleObjectBinaryPart,
)
from docx4j_py.openpackaging.parts.default_xml_part import (  # noqa: E402
    CustomXmlDataStoragePart,
    DefaultXmlPart,
)

SAMPLE = ROOT / "samples" / "toc.docx"


@pytest.fixture(scope="module")
def ctm() -> ContentTypeManager:
    with zipfile.ZipFile(SAMPLE) as zf:
        return ContentTypeManager.parse(zf.read("[Content_Types].xml"))


# ---------------------------------------------------------------------------
# parsing and lookup
# ---------------------------------------------------------------------------


def test_defaults_and_overrides_are_read(ctm):
    assert ctm.get_default_content_type("rels") == ContentTypes.RELATIONSHIPS_PART
    assert ctm.get_default_content_type("xml") == ContentTypes.APPLICATION_XML
    assert (
        ctm.get_override_content_type("/word/document.xml")
        == ContentTypes.WORDPROCESSINGML_DOCUMENT
    )


def test_an_override_beats_the_default(ctm):
    """``/word/document.xml`` would be ``application/xml`` by extension."""
    assert ctm.get_content_type("/word/document.xml") != ContentTypes.APPLICATION_XML
    assert ctm.get_content_type("/customXml/item1.xml") == ContentTypes.APPLICATION_XML


def test_lookup_is_case_insensitive(ctm):
    assert ctm.get_content_type("/WORD/DOCUMENT.XML") == ctm.get_content_type("/word/document.xml")
    assert ctm.get_default_content_type("XML") == ContentTypes.APPLICATION_XML


def test_an_unknown_extension_has_no_content_type(ctm):
    assert ctm.get_content_type("/word/thing.zzz") is None


def test_a_bad_root_is_refused():
    with pytest.raises(InvalidFormatException, match="expected Types"):
        ContentTypeManager.parse(b'<Nope xmlns="%s"/>' % CONTENT_TYPES_NS.encode())


# ---------------------------------------------------------------------------
# writing
# ---------------------------------------------------------------------------


def test_round_trip_of_the_file_itself(ctm):
    """Re-parsing what we write gives the same tables."""
    again = ContentTypeManager.parse(ctm.to_bytes())
    assert again.defaults == ctm.defaults
    assert again.overrides == ctm.overrides


def test_what_is_written_is_word_shaped(ctm):
    data = ctm.to_bytes()
    assert data.startswith(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    root = etree.fromstring(data[data.index(b"<Types") :])
    assert root.tag == f"{{{CONTENT_TYPES_NS}}}Types"
    assert root.nsmap == {None: CONTENT_TYPES_NS}
    locals_ = [etree.QName(c).localname for c in root]
    # defaults first, as Word writes them
    assert locals_ == sorted(locals_, key=lambda x: 0 if x == "Default" else 1)


def test_create_default_has_the_two_every_package_needs():
    ctm = ContentTypeManager.create_default()
    assert ctm.get_default_content_type("rels") == ContentTypes.RELATIONSHIPS_PART
    assert ctm.get_default_content_type("xml") == ContentTypes.APPLICATION_XML


def test_add_content_type_prefers_an_existing_default():
    ctm = ContentTypeManager.create_default()
    ctm.add_default_content_type("png", ContentTypes.IMAGE_PNG)
    ctm.add_content_type("/word/media/image1.png", ContentTypes.IMAGE_PNG)
    assert ctm.get_override_content_type("/word/media/image1.png") is None
    ctm.add_content_type("/word/media/image2.gif", ContentTypes.IMAGE_GIF)
    assert ctm.get_override_content_type("/word/media/image2.gif") == ContentTypes.IMAGE_GIF


def test_remove_content_type():
    ctm = ContentTypeManager.create_default()
    ctm.add_override_content_type("/word/document.xml", ContentTypes.WORDPROCESSINGML_DOCUMENT)
    ctm.remove_content_type("/word/document.xml")
    assert ctm.get_override_content_type("/word/document.xml") is None
    # with no override, it removes the extension's default
    ctm.remove_content_type("/word/other.xml")
    assert ctm.get_default_content_type("xml") is None


def test_part_names_for_a_content_type(ctm):
    names = ctm.get_part_names_for_content_type(ContentTypes.WORDPROCESSINGML_FOOTER)
    assert sorted(n.name for n in names) == ["/word/footer1.xml", "/word/footer2.xml"]


def test_is_content_type_registered(ctm):
    assert ctm.is_content_type_registered(ContentTypes.WORDPROCESSINGML_DOCUMENT)
    assert not ctm.is_content_type_registered("application/x-nonsense")


# ---------------------------------------------------------------------------
# the helpers
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("content_type", "expected"),
    [
        (ContentTypes.WORDPROCESSINGML_DOCUMENT, True),
        (ContentTypes.APPLICATION_XML, True),
        (ContentTypes.XML, True),
        (ContentTypes.IMAGE_PNG, False),
        (ContentTypes.VML_DRAWING, False),
        (None, False),
    ],
)
def test_is_xml_content_type(content_type, expected):
    assert is_xml_content_type(content_type) is expected


@pytest.mark.parametrize(
    ("content_type", "expected"),
    [
        (ContentTypes.IMAGE_PNG, True),
        (ContentTypes.IMAGE_JPEG, True),
        (ContentTypes.IMAGE_GIF, True),
        (ContentTypes.IMAGE_EMF, False),  # a metafile deflates well
        (ContentTypes.IMAGE_WMF, False),
        (ContentTypes.WORDPROCESSINGML_DOCUMENT, False),
    ],
)
def test_is_stored_uncompressed(content_type, expected):
    assert is_stored_uncompressed(content_type) is expected


# ---------------------------------------------------------------------------
# the registry, docx4j's newPartForContentType
# ---------------------------------------------------------------------------


class _Rel:
    """The smallest thing the registry reads a relationship type off."""

    def __init__(self, type_value: str) -> None:
        self.type_value = type_value


@pytest.mark.parametrize(
    ("name", "content_type", "cls"),
    [
        (
            "/word/document.xml",
            ContentTypes.WORDPROCESSINGML_DOCUMENT,
            MainDocumentPart,
        ),
        (
            "/word/styles.xml",
            ContentTypes.WORDPROCESSINGML_STYLES,
            StyleDefinitionsPart,
        ),
        (
            "/word/stylesWithEffects.xml",
            ContentTypes.WORDPROCESSINGML_STYLESWITHEFFECTS,
            StyleDefinitionsPart,
        ),
        ("/word/theme/theme1.xml", ContentTypes.OFFICEDOCUMENT_THEME, ThemePart),
        ("/word/media/image1.png", ContentTypes.IMAGE_PNG, ImagePart),
        ("/word/media/image1.tiff", ContentTypes.IMAGE_TIFF, ImagePart),
        (
            "/word/fonts/font1.odttf",
            ContentTypes.OFFICEDOCUMENT_FONT,
            ObfuscatedFontPart,
        ),
        ("/word/thing.bin", ContentTypes.OCTET_STREAM, BinaryPart),
        ("/word/unknown.xml", "application/x-unknown+xml", DefaultXmlPart),
        ("/word/unknown.zzz", "application/x-unknown", BinaryPart),
    ],
)
def test_the_content_type_table(name, content_type, cls):
    part = default_part_registry.create_part(name, content_type)
    assert type(part) is cls
    assert part.content_type == content_type


@pytest.mark.parametrize(
    ("relationship_type", "cls"),
    [
        (Namespaces.AF, AlternativeFormatInputPart),
        (Namespaces.EMBEDDED_PKG, EmbeddedPackagePart),
        (Namespaces.OLE_OBJECT, OleObjectBinaryPart),
        (Namespaces.CUSTOM_XML_DATA_STORAGE, CustomXmlDataStoragePart),
    ],
)
def test_the_relationship_type_wins_for_the_generic_content_types(relationship_type, cls):
    """docx4j tests these four before the content type, and so does this."""
    part = default_part_registry.create_part(
        "/word/thing.xml", ContentTypes.APPLICATION_XML, _Rel(relationship_type)
    )
    assert type(part) is cls


def test_application_xml_is_a_custom_xml_data_part():
    """As in docx4j: the two constants are the same string and it tests custom XML first."""
    part = default_part_registry.create_part("/customXml/item1.xml", ContentTypes.APPLICATION_XML)
    assert type(part) is CustomXmlDataStoragePart


def test_no_content_type_still_keeps_the_bytes():
    part = default_part_registry.create_part("/word/media/x.png", None, _Rel(Namespaces.IMAGE))
    assert type(part) is ImagePart
    assert part.content_type == ContentTypes.IMAGE_PNG
    assert type(default_part_registry.create_part("/word/x.bin", None)) is BinaryPart
    assert type(default_part_registry.create_part("/word/x.xml", None)) is DefaultXmlPart


def test_a_registry_can_be_extended_without_touching_the_default():
    class MyPart(BinaryPart):
        pass

    mine = default_part_registry.copy().register("application/x-mine", lambda n, ct: MyPart(n, ct))
    assert type(mine.create_part("/x.mine", "application/x-mine")) is MyPart
    assert type(default_part_registry.create_part("/x.mine", "application/x-mine")) is BinaryPart


def test_every_registered_factory_builds_a_part():
    """A smoke test over the whole table: nothing in it is broken."""
    for content_type in list(default_part_registry._by_content_type):  # noqa: SLF001
        part = default_part_registry.create_part("/word/thing.xml", content_type)
        assert part.part_name == PartName.of("/word/thing.xml")
