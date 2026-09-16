"""The error hierarchy, and the promise that an untouched part is untouched.

CR-003 section 3.1 (one hierarchy, every message says what to do instead),
section 10.4 (``BuilderError`` re-rooted, its codes unchanged) and section 7
("untouched parts still byte-identical through the content API").
"""

from __future__ import annotations

import sys

import pytest
from conftest import ROOT, SAMPLES, part_bytes, sample

sys.path.insert(0, str(ROOT / "scripts"))

import canon

from docx4j_py import Docx4JException, WordprocessingMLPackage
from docx4j_py.model.content import (
    ContentError,
    Docx4JError,
    InvalidTargetError,
    SpanError,
    StyleError,
)
from docx4j_py.model.content.errors import AddressError, BuilderError
from docx4j_py.wml import wml


def test_the_hierarchy_has_one_root():
    assert Docx4JError is Docx4JException
    for error in (ContentError, AddressError, InvalidTargetError, StyleError, SpanError):
        assert issubclass(error, Docx4JException)
    assert issubclass(AddressError, ContentError)
    assert issubclass(BuilderError, ContentError)
    assert issubclass(BuilderError, ValueError), "Phase A's call sites keep working"


def test_every_error_carries_a_code_a_message_and_a_hint():
    error = ContentError("something is wrong", code="test.code", hint="do this instead")
    assert error.code == "test.code"
    assert error.message == "something is wrong"
    assert error.hint == "do this instead"
    assert str(error) == "something is wrong (do this instead)"
    assert error.to_dict() == {
        "code": "test.code",
        "message": "something is wrong",
        "hint": "do this instead",
    }


def test_the_builder_error_codes_did_not_change():
    from docx4j_py.wml import r, rpr_from_elements

    # from the module, not the package: ``docx4j_py.wml.sdt`` is the module's
    # name as well as the builder's, so the package attribute is whichever was
    # bound first (a wart of the lazy re-export, recorded in CR-003 section 11)
    from docx4j_py.wml.sdt import sdt

    with pytest.raises(BuilderError) as raised:
        rpr_from_elements([("nonsense", None)])
    assert raised.value.code == "rpr.unknown_property"

    with pytest.raises(BuilderError) as raised:
        sdt([r("a")], kind="RepeatingSection")
    assert raised.value.code == "sdt.form_mismatch"

    # and it is still a ValueError, as CR-003 section 10.4 requires
    with pytest.raises(ValueError):
        rpr_from_elements([("nonsense", None)])
    assert BuilderError is __import__(
        "docx4j_py.wml.builders", fromlist=["BuilderError"]
    ).BuilderError


def test_a_run_at_body_level_is_refused_by_name(new_package):
    body = new_package.body
    run = wml("<w:r><w:t>x</w:t></w:r>")

    with pytest.raises(InvalidTargetError) as raised:
        body.insert_element(run)

    message = str(raised.value)
    assert raised.value.code == "target.invalid"
    assert "w:body cannot hold w:r (R)" in message, "what was passed"
    assert "it takes" in message and "w:p" in message and "w:tbl" in message, "what it takes"
    assert "paragraph.insert_xml()" in raised.value.hint


def test_the_other_refusals(new_package):
    body = new_package.body
    paragraph = body.insert_paragraph("x")

    with pytest.raises(ContentError) as raised:
        body.insert_element(paragraph.element, location="After")
    assert raised.value.code == "target.missing"

    with pytest.raises(ContentError) as raised:
        body.insert_element([])
    assert raised.value.code == "insert.empty"

    with pytest.raises(ContentError) as raised:
        body.insert_element(paragraph.element, location="Replace")
    assert raised.value.code == "location.invalid"

    with pytest.raises(ContentError) as raised:
        paragraph.insert_text("x", location="Before")
    assert raised.value.code == "location.invalid"


def test_the_style_error_lists_the_five_closest_names():
    package = sample("toc.docx")
    package.style_definitions_part.contents  # noqa: B018 - unmarshal it

    with pytest.raises(StyleError) as raised:
        package.body.paragraphs[0].style = "Headding Wun"
    assert raised.value.code == "style.not_found"
    assert str(raised.value).count("'") >= 10, "five names, each quoted"


@pytest.mark.parametrize(
    "name", ["sample-docx.docx", "toc.docx", "tables.docx", "2010-sample1.docx"]
)
def test_editing_the_body_leaves_every_other_part_byte_identical(name):
    before = part_bytes((SAMPLES / name).read_bytes())

    package = sample(name)
    body = package.body
    body.insert_paragraph("added by the content API")
    body.paragraphs[0].insert_text("edited ", location="Start")
    after = part_bytes(package.save())

    assert set(before) == set(after)
    changed = [entry for entry in before if before[entry] != after[entry]]
    assert changed == ["word/document.xml"], f"only the document changed, not {changed}"

    # and the edit really is there, on reload
    back = WordprocessingMLPackage.load(package.save())
    assert back.body.paragraphs[0].text.startswith("edited ")
    assert back.body.paragraphs[-1].text == "added by the content API"


def test_reading_the_body_re_marshals_only_the_part_it_reads():
    name = "sample-docx.docx"
    before = part_bytes((SAMPLES / name).read_bytes())

    package = sample(name)
    body = package.body
    assert body.text
    assert [p.style for p in body.paragraphs]
    assert body.search("the")
    assert not package.style_definitions_part.is_unmarshalled, "reading a style name does not"
    after = part_bytes(package.save())

    changed = [entry for entry in before if before[entry] != after[entry]]
    assert changed == ["word/document.xml"], f"only the part read was touched, not {changed}"
    # and that one round-trips canonically identical, which is CR-002's promise
    assert canon.normalize(before["word/document.xml"]) == canon.normalize(
        after["word/document.xml"]
    )
