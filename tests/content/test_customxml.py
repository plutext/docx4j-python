"""Custom XML parts, nodes and XPath. CR-003 section 3.7, Phase E.

``samples/invoice2013.docx`` is the fixture section 7 names: one custom XML
part, ``/customXml/item1.xml``, whose ``<invoice>`` has **no namespace** --- so
the twenty bindings Word wrote carry an empty ``w:prefixMappings`` and fully
positional paths such as ``/invoice[1]/customer[1]/contact[1]``. A namespaced
part is built here, because nothing in ``samples/`` carries one.

The promise this file exists for: **reading unmarshals nothing**. A document
whose custom XML is read from end to end, and then saved, is byte for byte what
it was.
"""

from __future__ import annotations

import pytest
from conftest import SAMPLES, part_bytes, reloaded, sample

from docx4j_py import WordprocessingMLPackage, create_package
from docx4j_py.model.content import BindingError
from docx4j_py.model.customxml import (
    CustomXmlNode,
    CustomXmlPart,
    canonical_xpath_of,
    format_prefix_mappings,
    parse_prefix_mappings,
)

#: The ``ds:itemID`` of ``samples/invoice2013.docx``'s one custom XML part.
ITEM_ID = "{5D7BA57F-1E52-4637-9F82-2D4025768D4F}"

#: A namespaced part, for the prefix-mapping cases the invoice cannot show.
GREETING = '<greeting xmlns="http://example.com/g"><to>World</to><to>Moon</to></greeting>'


@pytest.fixture(scope="module")
def invoice():
    """``samples/invoice2013.docx``, loaded once for the reads."""
    return sample("invoice2013.docx")


@pytest.fixture
def greeting_package():
    """A document this engine built, with one namespaced custom XML part."""
    package = create_package()
    package.id_seed = 20260917
    package.body.insert_paragraph("Hello")
    package.custom_xml_parts.add(GREETING)
    return package


# ---------------------------------------------------------------------------
# the collection
# ---------------------------------------------------------------------------


def test_the_collection_lists_the_parts_with_their_ids_and_namespaces(invoice):
    parts = invoice.custom_xml_parts

    assert len(parts) == 1
    part = parts.items[0]
    assert isinstance(part, CustomXmlPart)
    assert part.id == ITEM_ID
    assert str(part.part.part_name) == "/customXml/item1.xml"
    # the invoice's <invoice> is in no namespace at all
    assert part.namespace_uri == ""
    assert part.built_in is False
    assert part.schema_collection == ()


def test_get_item_ignores_case_and_braces_both_ways(invoice):
    parts = invoice.custom_xml_parts

    assert parts.get_item(ITEM_ID) is parts.items[0]
    assert parts.get_item(ITEM_ID.lower()) is parts.items[0]
    assert parts.get_item(ITEM_ID.strip("{}")) is parts.items[0]
    assert parts.get_item(ITEM_ID.strip("{}").lower()) is parts.items[0]
    assert parts.get_item("{00000000-0000-0000-0000-000000000000}") is None


def test_get_by_namespace_finds_a_namespaced_part(greeting_package):
    parts = greeting_package.custom_xml_parts

    assert [part.id for part in parts.get_by_namespace("http://example.com/g")] == [
        parts.items[0].id
    ]
    assert parts.get_by_namespace("urn:nothing") == []


def test_a_document_with_no_custom_xml_has_an_empty_collection(new_package):
    assert len(new_package.custom_xml_parts) == 0
    assert list(new_package.custom_xml_parts) == []
    assert new_package.custom_xml_parts.get_item(ITEM_ID) is None


# ---------------------------------------------------------------------------
# the nodes
# ---------------------------------------------------------------------------


def test_the_document_element_and_its_children(invoice):
    root = invoice.custom_xml_parts.items[0].document_element

    assert isinstance(root, CustomXmlNode)
    assert root.base_name == "invoice"
    assert root.namespace_uri == ""
    assert root.node_type == "Element"
    assert root.has_child_nodes()
    assert [child.base_name for child in root.child_elements][:4] == [
        "logo",
        "invoicenumber",
        "invoicedate",
        "customer",
    ]
    assert root.parent_node is None
    assert invoice.custom_xml_parts.items[0].get_xml().startswith("<invoice>")


def test_select_nodes_reads_known_values(invoice):
    part = invoice.custom_xml_parts.items[0]

    assert part.select_single_node("/invoice[1]/customer[1]/company[1]").text == "Contozo Inc"
    assert part.select_single_node("/invoice[1]/customer[1]/contact[1]").text == "John Citizen"
    assert part.select_single_node("/invoice[1]/invoicenumber[1]").text == "1234"
    assert len(part.select_nodes("/invoice/lines/lineitem")) == 2
    assert part.select_single_node("/invoice/nothing") is None
    # an attribute, and a relative XPath from a node
    applies = part.select_single_node("/invoice[1]/VAT[1]/@applies")
    assert applies.node_type == "Attribute"
    assert applies.base_name == "applies"
    assert applies.text == "true"
    customer = part.select_single_node("/invoice/customer")
    assert customer.select_single_node("company").text == "Contozo Inc"


def test_a_node_reports_words_canonical_xpath(invoice, greeting_package):
    part = invoice.custom_xml_parts.items[0]
    node = part.select_single_node("/invoice/customer/company")

    # fully positional, and no prefixes at all: the part has no namespace
    assert node.xpath == "/invoice[1]/customer[1]/company[1]"
    assert node.prefix_mappings == ""
    assert node.parent_node.base_name == "customer"
    assert node.owner_part is part

    applies = part.select_single_node("/invoice/VAT/@applies")
    assert applies.xpath == "/invoice[1]/VAT[1]/@applies"

    # a namespaced part numbers the prefixes ns0, ns1, ... as Word does
    other = greeting_package.custom_xml_parts.items[0]
    second = other.select_nodes("//ns0:to", "xmlns:ns0='http://example.com/g'")[1]
    assert second.xpath == "/ns0:greeting[1]/ns0:to[2]"
    assert second.prefix_mappings == "xmlns:ns0='http://example.com/g'"
    assert second.text == "Moon"


def test_the_prefix_mapping_string_parses_and_formats():
    parsed = parse_prefix_mappings("xmlns:ns0='http://a' xmlns:ns1=\"http://b\"")

    assert parsed == {"ns0": "http://a", "ns1": "http://b"}
    assert format_prefix_mappings(parsed) == "xmlns:ns0='http://a' xmlns:ns1='http://b'"
    assert parse_prefix_mappings(None) == {}
    assert parse_prefix_mappings("") == {}
    assert parse_prefix_mappings("xmlns='http://d'") == {"": "http://d"}


def test_a_caller_may_choose_any_prefix_for_a_selection(greeting_package):
    part = greeting_package.custom_xml_parts.items[0]

    chosen = part.select_nodes("/g:greeting/g:to", "xmlns:g='http://example.com/g'")
    assert [node.text for node in chosen] == ["World", "Moon"]
    # but the canonical path always starts numbering at ns0
    assert canonical_xpath_of(chosen[0].element).xpath == "/ns0:greeting[1]/ns0:to[1]"
    # and the part's own namespace manager is what None means
    assert part.namespace_manager.lookup_namespace("ns0") == "http://example.com/g"
    assert part.namespace_manager.lookup_prefix("http://example.com/g") == "ns0"
    assert [mapping.prefix for mapping in part.namespace_manager.items] == ["ns0"]
    assert [node.text for node in part.select_nodes("/ns0:greeting/ns0:to")] == ["World", "Moon"]


def test_a_bad_xpath_names_itself(invoice):
    with pytest.raises(BindingError) as error:
        invoice.custom_xml_parts.items[0].select_nodes("/invoice[[")

    assert error.value.code == "binding.bad_xpath"
    assert "/invoice[[" in error.value.message


# ---------------------------------------------------------------------------
# reading changes nothing
# ---------------------------------------------------------------------------


def test_reading_every_part_and_node_leaves_the_document_byte_for_byte():
    source = (SAMPLES / "invoice2013.docx").read_bytes()
    package = WordprocessingMLPackage.load(source)

    for part in package.custom_xml_parts:
        part.get_xml()
        assert part.namespace_uri == ""
        root = part.document_element
        for node in root.child_elements:
            node.text  # noqa: B018 - the read is the point
            for attribute in node.attributes:
                attribute.text  # noqa: B018
        assert part.is_modified is False
    package.custom_xml_parts.describe()

    before = part_bytes(source)
    after = part_bytes(package.save())
    for name in ("customXml/item1.xml", "customXml/itemProps1.xml", "word/document.xml"):
        assert after[name] == before[name], name
    assert after == before


def test_a_node_mutation_marks_the_part_and_only_that_part():
    source = (SAMPLES / "invoice2013.docx").read_bytes()
    package = WordprocessingMLPackage.load(source)
    part = package.custom_xml_parts.get_item(ITEM_ID)

    part.select_single_node("/invoice/customer/company").text = "Acme Ltd"

    assert part.is_modified is True
    after = part_bytes(package.save())
    before = part_bytes(source)
    assert after["customXml/item1.xml"] != before["customXml/item1.xml"]
    assert b"Acme Ltd" in after["customXml/item1.xml"]
    # the properties part and the main part are untouched
    assert after["customXml/itemProps1.xml"] == before["customXml/itemProps1.xml"]
    assert after["word/document.xml"] == before["word/document.xml"]


# ---------------------------------------------------------------------------
# editing the tree
# ---------------------------------------------------------------------------


def test_nodes_are_added_replaced_and_removed(greeting_package):
    part = greeting_package.custom_xml_parts.items[0]
    root = part.document_element

    root.append_child_node('<from xmlns="http://example.com/g">Ada</from>')
    root.append_child_node("note", "http://example.com/g", "Element", "hi")
    assert [node.base_name for node in root.child_elements] == ["to", "to", "from", "note"]
    assert root.child_elements[3].text == "hi"

    root.insert_node_before(
        '<first xmlns="http://example.com/g">x</first>', root.child_elements[0]
    )
    assert next(node.base_name for node in root.child_elements) == "first"

    root.replace_child_node(root.child_elements[0], '<first xmlns="http://example.com/g">y</first>')
    assert root.child_elements[0].text == "y"

    root.child_elements[0].delete()
    assert [node.base_name for node in root.child_elements] == ["to", "to", "from", "note"]

    root.remove_child(root.child_elements[3])
    assert [node.base_name for node in root.child_elements] == ["to", "to", "from"]

    package = reloaded(greeting_package)
    saved = package.custom_xml_parts.items[0]
    assert [node.base_name for node in saved.document_element.child_elements] == [
        "to",
        "to",
        "from",
    ]
    assert "<from>Ada</from>" in saved.get_xml() or "Ada" in saved.get_xml()


def test_the_document_element_cannot_be_deleted(greeting_package):
    with pytest.raises(BindingError) as error:
        greeting_package.custom_xml_parts.items[0].document_element.delete()

    assert error.value.code == "binding.document_element"


def test_the_part_level_xpath_editors(greeting_package):
    part = greeting_package.custom_xml_parts.items[0]
    mappings = "xmlns:ns0='http://example.com/g'"

    part.insert_element("/ns0:greeting", '<to xmlns="http://example.com/g">Mars</to>', mappings)
    assert [node.text for node in part.select_nodes("/ns0:greeting/ns0:to", mappings)] == [
        "World",
        "Moon",
        "Mars",
    ]

    part.update_element(
        "/ns0:greeting/ns0:to[1]", '<to xmlns="http://example.com/g">Earth</to>', mappings
    )
    part.delete_element("/ns0:greeting/ns0:to[2]", mappings)
    assert [node.text for node in part.select_nodes("/ns0:greeting/ns0:to", mappings)] == [
        "Earth",
        "Mars",
    ]

    part.insert_attribute("/ns0:greeting", "status", "draft", mappings)
    assert part.document_element.element.get("status") == "draft"
    part.update_attribute("/ns0:greeting", "status", "final", mappings)
    assert part.document_element.element.get("status") == "final"
    part.delete_attribute("/ns0:greeting", "status", mappings)
    assert part.document_element.element.get("status") is None

    with pytest.raises(BindingError) as error:
        part.delete_element("/ns0:greeting/ns0:nothing", mappings)
    assert error.value.code == "binding.no_match"


def test_set_xml_replaces_the_whole_part_and_survives_a_reload(greeting_package):
    part = greeting_package.custom_xml_parts.items[0]

    part.set_xml('<greeting xmlns="http://example.com/g"><to>Everyone</to></greeting>')

    assert part.select_single_node("/ns0:greeting/ns0:to").text == "Everyone"
    saved = reloaded(greeting_package).custom_xml_parts.items[0]
    assert saved.select_single_node("/ns0:greeting/ns0:to").text == "Everyone"

    with pytest.raises(BindingError) as error:
        part.set_xml("<not well formed")
    assert error.value.code == "binding.not_well_formed"


# ---------------------------------------------------------------------------
# add() and delete()
# ---------------------------------------------------------------------------


def test_add_writes_the_part_its_properties_part_and_the_relationship(greeting_package):
    part = greeting_package.custom_xml_parts.items[0]

    assert str(part.part.part_name) == "/customXml/item1.xml"
    assert part.id.startswith("{") and part.id.endswith("}")
    assert part.id == part.id.upper()
    assert len(part.id) == 38
    assert part.namespace_uri == "http://example.com/g"
    assert part.schema_collection == ("http://example.com/g",)

    props = greeting_package.get_part("/customXml/itemProps1.xml")
    assert props is not None
    assert part.id.encode() in props.xml
    # Word drops a custom XML part that only the package relates to, so the
    # relationship is on the main document part (docx4j's addPropertiesPart)
    main = greeting_package.main_document_part
    targets = [rel.target for rel in main.relationships_part.list]
    assert any("customXml/item1.xml" in target for target in targets)

    saved = reloaded(greeting_package)
    assert saved.custom_xml_parts.get_item(part.id) is not None


def test_add_refuses_xml_with_no_document_element(new_package):
    with pytest.raises(BindingError) as error:
        new_package.custom_xml_parts.add("not xml at all")

    assert error.value.code == "binding.not_well_formed"


def test_a_second_part_takes_the_next_free_name(greeting_package):
    second = greeting_package.custom_xml_parts.add("<other><a>1</a></other>")

    assert str(second.part.part_name) == "/customXml/item2.xml"
    assert greeting_package.get_part("/customXml/itemProps2.xml") is not None
    assert second.id != greeting_package.custom_xml_parts.items[0].id
    assert len(greeting_package.custom_xml_parts) == 2
    assert second.namespace_uri == ""
    assert second.schema_collection == ()


def test_delete_removes_both_parts_and_unlinks_the_mappings(greeting_package):
    control = greeting_package.body.paragraphs[0].insert_content_control("PlainText")
    part = greeting_package.custom_xml_parts.items[0]
    assert control.xml_mapping.set_mapping(
        "/ns0:greeting[1]/ns0:to[1]", "xmlns:ns0='http://example.com/g'", part
    )

    part.delete()

    assert greeting_package.get_part("/customXml/item1.xml") is None
    assert greeting_package.get_part("/customXml/itemProps1.xml") is None
    assert len(greeting_package.custom_xml_parts) == 0
    assert control.xml_mapping.is_mapped is False
    assert part.unlinked == ("/ns0:greeting[1]/ns0:to[1]",)
    # and the saved document has neither part nor binding
    saved = reloaded(greeting_package)
    assert "customXml" not in saved.body.content_controls[0].get_xml()
    assert len(saved.custom_xml_parts) == 0


def test_a_dry_run_undoes_an_added_part_and_leaves_the_document_alone():
    package = sample("invoice2013.docx")
    package.id_seed = 20260917
    before = package.save()

    with package.dry_run() as trial:
        added = trial.custom_xml_parts.add('<x xmlns="urn:x"><a>1</a></x>')
        assert str(added.part.part_name) == "/customXml/item2.xml"
        trial.custom_xml_parts.get_item(added.id).select_single_node(
            "/ns0:x/ns0:a", "xmlns:ns0='urn:x'"
        ).text = "2"

    assert package.get_part("/customXml/item2.xml") is None
    assert len(package.custom_xml_parts) == 1
    assert package.save() == before


def test_a_dry_run_refuses_to_delete_a_part():
    package = sample("invoice2013.docx")

    with package.dry_run() as trial, pytest.raises(BindingError.__mro__[1]) as error:
        trial.custom_xml_parts.items[0].delete()

    assert error.value.code == "dry_run.delete_part"
    assert len(package.custom_xml_parts) == 1
