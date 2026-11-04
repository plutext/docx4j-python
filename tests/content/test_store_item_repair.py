"""CR-014 phase 1c (section 2.10, Q11): a store item ID that names no part, resolved as Word repairs it.

docx4j CR-058 phase 2, ``StoreItemResolver``, ``35cdeaee6``. The oracle is docx4j's
``StoreItemRepairTest`` and its three fixtures, copied with a SHA-1 manifest
(``tests/fixtures/opendope-cr058/``); each test here says which of theirs it ports.
Word 365 ran the fixtures (docx4j CR-058 section 8): B7 showed its XML value and
Word's save carried the part's ID; of two parts that fit, the first in part order
whose XPath selects a node; where none fits, the control keeps its surface and its
stale ID.
"""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest
from conftest import reloaded

from docx4j_py import create_package, load
from docx4j_py.model.customxml.store_items import (
    first_step,
    resolve,
    resolved_store_item_id,
)
from docx4j_py.model.text import extract, plain_text
from docx4j_py.model.text.bound_values import BoundValues
from docx4j_py.wml import el
from docx4j_py.wml.sdt import sdt_property

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "opendope-cr058"
PART1 = "{B1D0F00D-0000-4000-8000-000000000001}"
PART2 = "{B1D0F00D-0000-4000-8000-000000000002}"
STALE = "{B1D0F00D-0000-4000-8000-0000000000FF}"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def fixture(name: str):
    return load(FIXTURES / name)


def text_of(package) -> str:
    return plain_text(extract(package))


def by_alias(package, alias: str):
    return next(control for control in package.body.content_controls if control.title == alias)


def shows(package, alias: str, value: str, store_item_id: str) -> None:
    control = by_alias(package, alias)
    assert control.text == value, alias
    assert control.xml_mapping.store_item_id == store_item_id, alias


def test_the_fixtures_are_docx4js_as_copied():
    for line in (FIXTURES / "SHA1SUMS").read_text().splitlines():
        digest, name = line.split()
        assert hashlib.sha1((FIXTURES / name).read_bytes()).hexdigest() == digest, name


# --- bind-plain B7: one part, which the XPath fits (Word's measured case) ---


def test_apply_bindings_resolves_a_stale_store_item_id_as_word_does():
    """``theGeneratorResolvesAStaleStoreItemIdAsWordDoes``."""
    package = fixture("bind-plain.docx")
    result = package.custom_xml_parts.apply_bindings()
    shows(package, "B1", "B1_XML_VALUE", PART1)
    shows(package, "B7", "B7_XML_VALUE", PART1)  # the ID written back, as Word's save has it
    entry = next(entry for entry in result.applied if entry.title == "B7")
    assert entry.store_item_id == PART1
    assert [note.code for note in entry.notes] == ["binding.store_item_repaired"]
    assert STALE in entry.notes[0].message and PART1 in entry.notes[0].message
    assert entry.to_dict()["notes"][0]["code"] == "binding.store_item_repaired"
    assert any(STALE in warning for warning in package.changes[-1].warnings)
    untouched = next(entry for entry in result.applied if entry.title == "B1")
    assert untouched.notes == ()


def test_the_repaired_id_is_saved():
    package = fixture("bind-plain.docx")
    package.custom_xml_parts.apply_bindings()
    again = reloaded(package)
    shows(again, "B7", "B7_XML_VALUE", PART1)
    assert by_alias(again, "B7").xml_mapping.custom_xml_part is not None


def test_the_text_model_shows_the_value_word_shows():
    """``theTextModelShowsTheValueWordShows``."""
    text = text_of(fixture("bind-plain.docx"))
    assert "B7_XML_VALUE" in text and "B1_XML_VALUE" in text
    assert "B7_STALE_SURFACE" not in text


def test_the_text_model_writes_nothing():
    package = fixture("bind-plain.docx")
    data = (FIXTURES / "bind-plain.docx").read_bytes()
    text_of(package)
    assert by_alias(load(data), "B7").xml_mapping.store_item_id == STALE
    assert by_alias(package, "B7").xml_mapping.store_item_id == STALE, "BoundValues writes nothing"


def test_update_from_content_controls_writes_the_part_word_would():
    package = fixture("bind-plain.docx")
    control = by_alias(package, "B7")
    part = package.custom_xml_parts.get_item(PART1)
    before = control.text
    result = package.custom_xml_parts.update_from_content_controls()
    entry = next(entry for entry in result.applied if entry.title == "B7")
    assert entry.value == before
    assert control.xml_mapping.store_item_id == STALE, "this direction only reads the document"
    assert part.select_single_node(control.xml_mapping.xpath, control.xml_mapping.prefix_mappings).text == before


def test_a_text_write_reaches_the_part_the_binding_resolves_to():
    package = fixture("bind-plain.docx")
    control = by_alias(package, "B7")
    control.insert_text("typed", location="Replace")
    part = package.custom_xml_parts.get_item(PART1)
    assert part.select_single_node(control.xml_mapping.xpath, control.xml_mapping.prefix_mappings).text == "typed"


def test_xml_mapping_stays_strict():
    """Q11: ``custom_xml_part`` is Office JS's read, and what Office JS gives for a stale ID is unmeasured."""
    control = by_alias(fixture("bind-plain.docx"), "B7")
    assert control.xml_mapping.custom_xml_part is None
    assert control.xml_mapping.custom_xml_node is None
    assert control.xml_mapping.resolved_part is not None
    assert control.xml_mapping.resolved_node.text == "B7_XML_VALUE"
    assert control.xml_mapping.to_dict()["resolves"] is False


# --- two parts fit ---


def test_where_two_parts_fit_the_first_that_yields_a_node_is_bound():
    """``whereTwoPartsFitTheFirstThatYieldsANodeIsBound``."""
    package = fixture("bind-plain-two-parts.docx")
    package.custom_xml_parts.apply_bindings()
    shows(package, "B1", "B1_XML_VALUE", PART1)
    shows(package, "B7", "B7_XML_VALUE_PART1", PART1)  # both yield: the first in part-name order
    shows(package, "B11", "B11_XML_VALUE_PART2", PART2)  # only part 2 yields
    text = text_of(fixture("bind-plain-two-parts.docx"))
    assert "B7_XML_VALUE_PART1" in text and "B11_XML_VALUE_PART2" in text


def test_where_two_parts_fit_and_neither_selects_a_node_the_first_is_named():
    package = fixture("bind-plain-two-parts.docx")
    collection = package.custom_xml_parts
    control = by_alias(package, "B7")
    mappings = control.xml_mapping.prefix_mappings
    part = resolve(collection, STALE, "/p:root[1]/p:nothing[1]", mappings)
    assert part is not None and part.id == PART1


# --- no part fits ---


def test_where_no_part_fits_the_control_is_left_and_the_rest_binds():
    """``whereNoPartFitsTheControlIsLeftAndTheRestBinds``."""
    package = fixture("bind-plain-no-part.docx")
    result = package.custom_xml_parts.apply_bindings()
    shows(package, "B1", "B1_XML_VALUE", PART1)
    shows(package, "B7", "B7_STALE_SURFACE", STALE)
    shows(package, "B12", "B12_STALE_SURFACE", STALE)
    assert {entry.title: entry.code for entry in result.skipped} == {"B7": "no_part", "B12": "no_part"}


def test_the_text_model_shows_the_surface_where_no_part_fits():
    """``theTextModelShowsTheSurfaceWhereNoPartFits``."""
    text = text_of(fixture("bind-plain-no-part.docx"))
    assert "B1_XML_VALUE" in text and "B7_STALE_SURFACE" in text and "B12_STALE_SURFACE" in text
    assert "B7_XML_VALUE" not in text


def test_update_from_content_controls_skips_where_no_part_fits():
    package = fixture("bind-plain-no-part.docx")
    result = package.custom_xml_parts.update_from_content_controls()
    assert {entry.code for entry in result.skipped if entry.title in ("B7", "B12")} == {"no_part"}


# --- the rule, as docx4j's firstStep reads it ---


@pytest.mark.parametrize(
    ("xpath", "mappings", "step"),
    [
        ("/p:root[1]/p:b7[1]", "xmlns:p='urn:x'", ("urn:x", "root")),
        ("/root[1]/b7[1]", "", ("", "root")),
        ("  /p:root/p:b7", "xmlns:p='urn:x'", ("urn:x", "root")),
        ("p:root", "xmlns:p='urn:x'", ("urn:x", "root")),
        ("//p:b7", "xmlns:p='urn:x'", None),
        ("/*[1]/p:b7[1]", "xmlns:p='urn:x'", None),
        ("/p:*[1]", "xmlns:p='urn:x'", None),
        ("/.", "", None),
        ("/..", "", None),
        ("/@x", "", None),
        ("/child::p:root", "xmlns:p='urn:x'", None),
        ("", "", None),
        (None, "", None),
        ("/q:root[1]", "xmlns:p='urn:x'", ("", "root")),
        ("/w:document[1]", "", (W, "document")),
    ],
)
def test_first_step(xpath, mappings, step):
    assert first_step(xpath, mappings) == step


def repair_package():
    package = create_package()
    package.body.insert_paragraph("seed")
    return package


def test_an_xpath_with_no_named_first_step_fits_every_part():
    package = repair_package()
    package.custom_xml_parts.add("<a><b>one</b></a>")
    second = package.custom_xml_parts.add("<a2><c>two</c></a2>")
    assert resolve(package.custom_xml_parts, STALE, "//c", "") == second


def test_the_name_part_order_is_docx4js():
    """``partName`` compared as text: ``item10`` before ``item2``. Word's rule is measured on two parts only."""
    package = repair_package()
    parts = [package.custom_xml_parts.add(f"<a><v>{index}</v></a>") for index in range(1, 11)]
    named = sorted(parts, key=lambda part: str(part.part.part_name))
    assert str(named[0].part.part_name) == "/customXml/item1.xml"
    assert str(named[1].part.part_name) == "/customXml/item10.xml"
    assert resolve(package.custom_xml_parts, STALE, "/a[1]/v[1]", "") == named[0]


def test_the_id_written_is_the_parts_upper_cased():
    package = repair_package()
    part = package.custom_xml_parts.add("<a><v>x</v></a>")
    assert resolved_store_item_id(package.custom_xml_parts, STALE, "/a[1]/v[1]", "") == part.id.upper()
    assert resolved_store_item_id(package.custom_xml_parts, part.id.lower(), "/a[1]/v[1]", "") == part.id.lower()
    assert resolved_store_item_id(package.custom_xml_parts, STALE, "/none[1]", "") == STALE


def test_a_properties_store_is_not_stale():
    """The well-known IDs name no part here by design (CR-003 section 4), and nothing is repaired over them."""
    package = repair_package()
    package.custom_xml_parts.add("<coreProperties><title>x</title></coreProperties>")
    core = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}"
    extended = "{6668398D-A668-4E3E-A5EB-62B293D839F1}"
    cover = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}"
    for store in (core, extended, cover, core.lower()):
        assert resolve(package.custom_xml_parts, store, "/coreProperties[1]/title[1]", "") is None


def test_a_word_property_store_part_is_no_candidate():
    """docx4j's ``instanceof CustomXmlDataStoragePart``: the cover-page properties are a typed part there."""
    package = repair_package()
    package.custom_xml_parts.add('<CoverPageProperties xmlns="http://schemas.microsoft.com/office/2006/coverPageProps"/>')
    assert resolve(package.custom_xml_parts, STALE, "//x", "") is None


def test_a_case_or_brace_difference_is_no_stale_id():
    package = fixture("bind-plain.docx")
    assert resolve(package.custom_xml_parts, PART1.lower().strip("{}"), "/x", "").id == PART1


def test_no_part_where_nothing_was_named():
    package = fixture("bind-plain.docx")
    assert resolve(package.custom_xml_parts, "", "/x", "") is None
    assert resolve(None, STALE, "/x", "") is None


# --- what a repair doesn't reach ---


def test_a_stale_binding_left_unapplied_keeps_its_id():
    """Word repairs on open, but a control the pass skips is left as it is (an open point of CR-014's)."""
    package = repair_package()
    package.custom_xml_parts.add("<a><v>x</v></a>")
    package.body.insert_xml(
        f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="7"/><w:picture/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>'
    )
    control = package.body.content_controls[0]
    control.put_property(el.dataBinding(xpath="/a[1]/v[1]", store_item_id=STALE))
    result = package.custom_xml_parts.apply_bindings()
    assert [entry.code for entry in result.skipped] == ["picture"]
    assert sdt_property(control.sdt_pr, "dataBinding").store_item_id == STALE


def test_a_fitting_part_with_no_node_is_no_match_and_keeps_the_id():
    """Word's "no part" probes have no root match; this shape is unmeasured, so nothing is changed."""
    package = repair_package()
    package.custom_xml_parts.add("<a><v>x</v></a>")
    package.body.insert_xml(
        f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>surface</w:t></w:r></w:p>'
        "</w:sdtContent></w:sdt>"
    )
    control = package.body.content_controls[0]
    control.put_property(el.dataBinding(xpath="/a[1]/none[1]", store_item_id=STALE))
    result = package.custom_xml_parts.apply_bindings()
    assert [entry.code for entry in result.skipped] == ["no_match"]
    assert control.text == "surface"
    assert control.xml_mapping.store_item_id == STALE


def test_a_dry_run_of_a_repair_leaves_the_real_id():
    package = fixture("bind-plain.docx")
    with package.dry_run() as trial:
        trial.custom_xml_parts.apply_bindings()
        assert by_alias(trial, "B7").xml_mapping.store_item_id == PART1
    assert by_alias(package, "B7").xml_mapping.store_item_id == STALE


def test_a_flat_opc_node_reached_through_a_stale_id_is_read_as_one():
    """:func:`~docx4j_py.model.content.controls.is_bound_block_control` reads the node Word reads."""
    package = repair_package()
    flat = '&lt;pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"/&gt;'
    package.custom_xml_parts.add(f"<a><v>{flat}</v></a>")
    package.body.insert_xml(
        f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>'
    )
    control = package.body.content_controls[0]
    control.put_property(el.dataBinding(xpath="/a[1]/v[1]", store_item_id=STALE))
    assert control.type == "RichText"


def test_bound_values_reads_each_xpath_of_a_stale_id_in_its_own_part():
    package = fixture("bind-plain-two-parts.docx")
    reader = BoundValues(package)
    b7 = sdt_property(by_alias(package, "B7").sdt_pr, "dataBinding")
    b11 = sdt_property(by_alias(package, "B11").sdt_pr, "dataBinding")
    assert reader.value(b7) == "B7_XML_VALUE_PART1"
    assert reader.value(b11) == "B11_XML_VALUE_PART2"
    assert reader.value(b7) == "B7_XML_VALUE_PART1"


# --- describe() and fill() read as the pass does (section 13.3, O1) ---


def test_describe_reads_a_stale_binding_from_the_part_its_xpath_fits():
    package = fixture("bind-plain.docx")
    b7 = next(binding for binding in package.custom_xml_parts.describe().bindings if binding.tag == "B7")
    assert b7.value == "B7_XML_VALUE"
    assert b7.store_item_id == STALE, "a read writes nothing, and reports the binding as it stands"


def test_fill_writes_a_stale_keys_node_and_the_pass_after_it_repairs_the_id():
    package = fixture("bind-plain.docx")
    package.custom_xml_parts.fill({"B7": "FILLED"})
    shows(package, "B7", "FILLED", PART1)


def test_fill_still_skips_a_key_no_part_fits():
    package = fixture("bind-plain-no-part.docx")
    result = package.custom_xml_parts.fill({"B7": "FILLED"})
    assert [entry.code for entry in result.skipped] == ["no_part"]
    assert by_alias(package, "B7").text == "B7_STALE_SURFACE"
