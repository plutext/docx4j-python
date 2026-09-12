"""The generated classes carry docx4j's names, CR-001 section 6.1.

`codegen/names/` maps every schema type and element name to the name docx4j's
XJC run gave it, and the fork's `<ClassNames>` option applies it. What matters
is that the classes a docx4j user already knows are there under those names,
and that renaming a class never changed the document it reads or writes: the
xml name lives in `Meta.name`, which is the schema's, not the table's.

    .venv-fork/bin/python -m pytest tests/test_names.py -q
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py import wml
from docx4j_py.child import Child

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# The names CR-001 section 6.1 names, and the xml element each one is for.
# None means the class is a schema type, not a global element: it has no xml
# name of its own and the field that holds it supplies one.
PART_ROOTS = {
    "Document": "document",
    "Styles": "styles",
    "Numbering": "numbering",
    "Fonts": "fonts",
    "Comments": "comments",
    "Hdr": "hdr",
    "Ftr": "ftr",
    "Footnotes": "footnotes",
    "Endnotes": "endnotes",
    "Settings": "settings",
    "WebSettings": "webSettings",
    "GlossaryDocument": "glossaryDocument",
}

ELEMENTS = {
    "P": "p",
    "R": "r",
    "Tr": "tr",
    "Body": "body",
    "Style": "style",
    "RunIns": "ins",
    "RunDel": "del",
    "DelText": "delText",
    "SdtBlock": "sdt",
    "Br": "br",
}

TYPES = [
    "PPr",
    "RPr",
    "Tbl",
    "Tc",
    "Drawing",
    "Pict",
    "Text",
    "CTSettings",
    "CTShd",
    "CTMarkup",
    "CTTrackChange",
    "RunTrackChange",
    "BooleanDefaultTrue",
    "TblWidth",
    "HpsMeasure",
    "CTBookmark",
    "CTLanguage",
    "CTColor",
    "SdtRun",
    "CTSdtCell",
    "CTSdtRow",
]

ENUMS = ["JcEnumeration", "STBrType", "STThemeColor"]


@pytest.mark.parametrize("name", sorted(set(PART_ROOTS) | set(ELEMENTS) | set(TYPES)))
def test_the_docx4j_class_exists(name: str) -> None:
    klass = getattr(wml, name)

    assert isinstance(klass, type)
    assert issubclass(klass, Child)


@pytest.mark.parametrize("name", ENUMS)
def test_the_docx4j_enum_exists(name: str) -> None:
    import enum

    assert issubclass(getattr(wml, name), enum.Enum)


@pytest.mark.parametrize(
    ("name", "xml_name"), sorted({**PART_ROOTS, **ELEMENTS}.items())
)
def test_meta_name_is_the_xml_name(name: str, xml_name: str) -> None:
    """A rename never changes what the class reads or writes."""
    meta = getattr(wml, name).Meta

    assert meta.name == xml_name
    assert meta.namespace == WML


def test_a_renamed_type_keeps_its_schema_name() -> None:
    """A type class's Meta.name is the schema type name, not the table's."""
    assert wml.Text.Meta.name == "CT_Text"
    assert wml.PPr.Meta.name == "CT_PPr"
    assert wml.CTSettings.Meta.name == "CT_Settings"
    # The docx4j schema's own jaxb:class rename, not stock ECMA-376.
    assert wml.JcEnumeration.__name__ == "JcEnumeration"


def test_the_name_table_did_not_break_the_run_content() -> None:
    """w:t in a w:r is a Text, whatever the class the generator made for it."""
    assert issubclass(wml.RT, wml.Text)
    assert issubclass(wml.RInstrText, wml.Text)
    assert issubclass(wml.RDelInstrText, wml.Text)
    assert issubclass(wml.DelText, Child)

    # isinstance over the shared type is what CR-001 section 11 question 2
    # promised in place of JAXBElement.
    assert isinstance(wml.RT(value="hello"), wml.Text)


def test_the_element_specific_track_changes_are_docx4j_s() -> None:
    assert issubclass(wml.RunIns, wml.CTTrackChange)
    assert issubclass(wml.RunDel, wml.CTTrackChange)
    assert wml.RunIns.Meta.name == "ins"
    assert wml.RunDel.Meta.name == "del"


def test_a_paragraph_built_by_name_serialises() -> None:
    from docx4j_xsdata.formats.dataclass.serializers import XmlSerializer
    from docx4j_xsdata.formats.dataclass.serializers.config import SerializerConfig

    p = wml.P(
        p_pr=wml.PPr(r_pr=wml.ParaRPr(b=wml.BooleanDefaultTrue())),
        content=[wml.R(content=[wml.RT(value="Hello")])],
    )
    config = SerializerConfig(xml_declaration=False, bool_format="numeric")
    xml = XmlSerializer(config=config).render(p, ns_map={"w": WML})

    assert xml == (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>"
        "<w:r><w:t>Hello</w:t></w:r></w:p>"
    )


def test_the_package_has_no_leftover_generator_names() -> None:
    """The names the table replaced are gone, not shadowed."""
    for gone in ("CtPpr", "CtRpr", "CtText", "CtSettings", "CtTbl", "CtP"):
        assert not hasattr(wml, gone), gone


def test_no_docx4j_name_was_pushed_aside_by_a_suffix() -> None:
    """The table must not spend a docx4j name on an invented class.

    A name table derived from JAXB maps many element names onto one payload
    class, which the naive application of it turned into `Text1` with
    `Text2`..`Text7` over the elements that share it (fork CHANGES.md stage 4).
    """
    for stolen in ("Text1", "BooleanDefaultTrue1", "CTMarkup1", "CTTrackChange1"):
        assert not hasattr(wml, stolen), stolen
