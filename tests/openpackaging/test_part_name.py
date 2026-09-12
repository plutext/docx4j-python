"""``PartName``: resolution, relativisation, validation, case-insensitivity.

CR-002 section 5.1 and decided question 6. Mirrors docx4j-core-ts's
``test/partname.test.mjs``.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_part_name.py -q
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py.openpackaging import InvalidFormatException, PartName  # noqa: E402

# ---------------------------------------------------------------------------
# construction and validation
# ---------------------------------------------------------------------------


def test_a_leading_slash_is_added():
    assert PartName.of("word/document.xml").name == "/word/document.xml"
    assert PartName.of("/word/document.xml").name == "/word/document.xml"


def test_of_is_idempotent():
    name = PartName.of("/word/document.xml")
    assert PartName.of(name) is name


@pytest.mark.parametrize(
    ("value", "clause"),
    [
        ("", "M1.1"),
        ("/", "M1.1"),
        ("/word/", "M1.5"),
        ("/word//document.xml", "M1.3"),
        ("/word/document.", "M1.9"),
        ("/word/../document.xml", "M1.9"),  # a segment ending in a dot is caught first
        ("/word/.../document.xml", "M1.9"),
        ("/word/doc ument.xml", "M1.6"),
        ("/word/a%2Fb.xml", "M1.7"),
    ],
)
def test_the_opc_conformance_clauses(value, clause):
    with pytest.raises(InvalidFormatException, match=clause):
        PartName.of(value)


def test_an_absolute_uri_is_not_a_part_name():
    with pytest.raises(InvalidFormatException, match="Absolute URI"):
        PartName.of("http://example.com/x.xml")


def test_a_bad_percent_escape_is_refused():
    with pytest.raises(InvalidFormatException, match="invalid encoded character"):
        PartName.of("/word/a%zz.xml")


def test_the_root_is_not_a_valid_part_name_but_exists():
    """docx4j builds it with ``checkConformance=false``; so does this."""
    assert PartName.ROOT.name == "/"
    with pytest.raises(InvalidFormatException):
        PartName.of("/")


# ---------------------------------------------------------------------------
# the pieces
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("name", "extension", "directory", "file_name", "store_name"),
    [
        ("/word/document.xml", "xml", "/word", "document.xml", "word/document.xml"),
        (
            "/word/media/image1.png",
            "png",
            "/word/media",
            "image1.png",
            "word/media/image1.png",
        ),
        ("/docProps/core.xml", "xml", "/docProps", "core.xml", "docProps/core.xml"),
        ("/noextension", "", "/", "noextension", "noextension"),
    ],
)
def test_the_pieces(name, extension, directory, file_name, store_name):
    pn = PartName.of(name)
    assert pn.extension == extension
    assert pn.directory == directory
    assert pn.file_name == file_name
    assert pn.store_name == store_name


def test_the_extension_uses_the_last_dot():
    """docx4j's ``getPart`` uses the *first*; the last is the right answer."""
    assert PartName.of("/word/my.file.name.png").extension == "png"


# ---------------------------------------------------------------------------
# case insensitivity, decided question 6
# ---------------------------------------------------------------------------


def test_equality_and_hashing_are_case_insensitive():
    a = PartName.of("/word/Header1.xml")
    b = PartName.of("/word/header1.xml")
    assert a == b
    assert hash(a) == hash(b)
    assert len({a, b}) == 1
    assert a == "/WORD/HEADER1.XML"


def test_the_stored_spelling_survives():
    """Lookup folds case; writing does not."""
    assert PartName.of("/word/Header1.xml").name == "/word/Header1.xml"


def test_ordering_is_by_the_folded_name():
    names = sorted(PartName.of(n) for n in ("/b.xml", "/A.xml", "/c.xml"))
    assert [n.name for n in names] == ["/A.xml", "/b.xml", "/c.xml"]


# ---------------------------------------------------------------------------
# relationships parts
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("source", "rels"),
    [
        ("/", "/_rels/.rels"),
        ("/word/document.xml", "/word/_rels/document.xml.rels"),
        ("/word/media/image1.png", "/word/media/_rels/image1.png.rels"),
        ("/docProps/core.xml", "/docProps/_rels/core.xml.rels"),
    ],
)
def test_rels_for_and_back(source, rels):
    assert PartName.rels_for(source).name == rels
    assert PartName.source_of_rels(rels).name == source


def test_source_of_rels_refuses_a_name_that_is_not_one():
    with pytest.raises(InvalidFormatException, match="Not a relationships part name"):
        PartName.source_of_rels("/word/document.xml")


def test_is_relationships_part():
    assert PartName.of("/_rels/.rels").is_relationships_part
    assert PartName.of("/word/_rels/document.xml.rels").is_relationships_part
    assert not PartName.of("/word/document.xml").is_relationships_part


# ---------------------------------------------------------------------------
# resolve and relativize
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("source", "target", "expected"),
    [
        ("/", "word/document.xml", "/word/document.xml"),
        ("/", "/word/document.xml", "/word/document.xml"),
        ("/word/document.xml", "styles.xml", "/word/styles.xml"),
        ("/word/document.xml", "media/image1.png", "/word/media/image1.png"),
        ("/word/document.xml", "../customXml/item1.xml", "/customXml/item1.xml"),
        ("/word/document.xml", "./theme/theme1.xml", "/word/theme/theme1.xml"),
        ("/word/glossary/document.xml", "../styles.xml", "/word/styles.xml"),
        ("/word/document.xml", "/docProps/core.xml", "/docProps/core.xml"),
    ],
)
def test_resolve(source, target, expected):
    assert PartName.resolve(source, target).name == expected


@pytest.mark.parametrize(
    ("source", "target", "expected"),
    [
        ("/", "/word/document.xml", "word/document.xml"),
        ("/word/document.xml", "/word/styles.xml", "styles.xml"),
        ("/word/document.xml", "/word/media/image1.png", "media/image1.png"),
        ("/word/document.xml", "/customXml/item1.xml", "../customXml/item1.xml"),
        ("/customXml/item1.xml", "/customXml/itemProps1.xml", "itemProps1.xml"),
    ],
)
def test_relativize(source, target, expected):
    assert PartName.relativize(source, target) == expected


@pytest.mark.parametrize(
    ("source", "target"),
    [
        ("/", "/word/document.xml"),
        ("/word/document.xml", "/word/styles.xml"),
        ("/word/document.xml", "/word/media/image1.png"),
        ("/word/document.xml", "/customXml/item1.xml"),
        ("/word/glossary/document.xml", "/word/styles.xml"),
    ],
)
def test_relativize_and_resolve_are_inverses(source, target):
    assert PartName.resolve(source, PartName.relativize(source, target)).name == target


def test_repr_and_str():
    pn = PartName.of("/word/document.xml")
    assert str(pn) == "/word/document.xml"
    assert repr(pn) == "PartName('/word/document.xml')"
