"""The template story, tool shaped: describe, fill, save, reload, read back.

CR-003 section 3.7 and section 7's agent scenarios, Phase E. These are the
library half of docx4j-mcp's ``describe_template`` and ``fill_template``: every
step is one call with JSON-ready arguments and a JSON-ready result, and what is
pinned is the surface a server is written over.
"""

from __future__ import annotations

import json

import pytest
from conftest import SEED, part_bytes, reloaded, sample

from docx4j_py import create_package
from docx4j_py.model.content import BindingError

ITEM_ID = "{5D7BA57F-1E52-4637-9F82-2D4025768D4F}"

#: How big a ``describe()`` of the invoice may get before it stops being a tool
#: result. CR-003 section 3.4's budgets, the same shape as the outline's.
DESCRIBE_BUDGET = 12_000


@pytest.fixture
def invoice():
    """``samples/invoice2013.docx``: twenty bindings over one custom XML part."""
    return sample("invoice2013.docx")


def by_title(package, title: str):
    """The control whose ``w:alias`` is `title`."""
    return next(control for control in package.body.content_controls if control.title == title)


# ---------------------------------------------------------------------------
# describe()
# ---------------------------------------------------------------------------


def test_describe_says_what_the_template_wants(invoice):
    skeleton = invoice.custom_xml_parts.describe()

    assert len(skeleton) == 20
    assert skeleton.resolved == 20
    assert skeleton.repeats == (
        "/invoice[1]/lines[1]/lineitem[1]",
        "/invoice[1]/notes[1]/note[1]",
    )
    assert [part.id for part in skeleton.parts] == [ITEM_ID]
    assert skeleton.parts[0].namespace_uri == ""
    assert skeleton.parts[0].built_in is False
    assert skeleton.open_dope is False

    kinds = {binding.kind for binding in skeleton.bindings}
    assert {"PlainText", "DatePicker", "CheckBox", "Picture", "RepeatingSection"} <= kinds

    company = next(
        binding
        for binding in skeleton.bindings
        if binding.xpath == "/invoice[1]/customer[1]/company[1]"
    )
    assert company.kind == "PlainText"
    assert company.value == "Contozo Inc"
    assert company.store_item_id == ITEM_ID
    assert company.part_name == "/customXml/item1.xml"
    assert company.repeat is None

    inside = next(
        binding
        for binding in skeleton.bindings
        if binding.xpath == "/invoice[1]/lines[1]/lineitem[1]/productcode[1]"
    )
    assert inside.repeat == "/invoice[1]/lines[1]/lineitem[1]"


def test_describe_reads_nothing_it_does_not_need(invoice):
    """The main document part is read as bytes unless it is already a tree."""
    before = part_bytes(invoice.save())

    skeleton = invoice.custom_xml_parts.describe()

    assert invoice.main_document_part.is_unmarshalled is False
    assert any("read as bytes" in warning for warning in skeleton.warnings)
    assert all(binding.address == "" for binding in skeleton.bindings)
    assert part_bytes(invoice.save()) == before


def test_describe_reports_addresses_once_the_body_has_been_read(invoice):
    invoice.body.paragraphs  # noqa: B018 - this is what unmarshals the part

    skeleton = invoice.custom_xml_parts.describe()

    assert skeleton.warnings == ()
    assert all(binding.address.startswith("body/") for binding in skeleton.bindings)
    assert len(skeleton) == 20


def test_the_describe_json_stays_inside_a_tool_budget(invoice):
    payload = invoice.custom_xml_parts.describe().to_json()

    assert len(payload) < DESCRIBE_BUDGET
    parsed = json.loads(payload)
    assert set(parsed) == {"bindings", "repeats", "parts", "open_dope", "resolved", "warnings"}
    assert len(parsed["bindings"]) == 20
    # the picture binding's node holds the whole image as base64; the skeleton
    # reports only the first 200 characters of it
    logo = next(entry for entry in parsed["bindings"] if entry["kind"] == "Picture")
    assert len(logo["value"]) == 201
    assert logo["value"].endswith("…")
    # and asking for all of it does give all of it
    full = invoice.custom_xml_parts.describe(max_value_chars=None)
    whole = next(binding for binding in full.bindings if binding.kind == "Picture")
    assert len(whole.value) > 9_000


# ---------------------------------------------------------------------------
# fill()
# ---------------------------------------------------------------------------


def test_fill_sets_the_nodes_and_the_controls_show_the_values(invoice):
    result = invoice.custom_xml_parts.fill(
        {
            "/invoice[1]/customer[1]/company[1]": "Acme Ltd",
            "/invoice[1]/invoicenumber[1]": "INV-9",
            "/invoice[1]/VAT[1]/@applies": "false",
            "/invoice[1]/invoicedate[1]": "2026-09-17T00:00:00Z",
        }
    )

    assert [entry.key for entry in result.applied] == [
        "/invoice[1]/customer[1]/company[1]",
        "/invoice[1]/invoicenumber[1]",
        "/invoice[1]/VAT[1]/@applies",
        "/invoice[1]/invoicedate[1]",
    ]
    assert result.skipped == ()
    assert result.bindings.updated == 16

    saved = reloaded(invoice)
    assert by_title(saved, "/invoice[1]/customer[1]/company[1]").text == "Acme Ltd"
    assert by_title(saved, "/invoice[1]/invoicenumber[1]").text == "INV-9"
    checkbox = next(c for c in saved.body.content_controls if c.type == "CheckBox")
    assert checkbox.checkbox_content_control.is_checked is False
    assert checkbox.text == "☐"
    date = next(c for c in saved.body.content_controls if c.type == "DatePicker")
    assert date.text == "17 September 2026"
    assert date.date_picker_content_control.full_date.year == 2026
    # the data itself carries the new values
    assert "Acme Ltd" in saved.custom_xml_parts.get_item(ITEM_ID).get_xml()


def test_fill_takes_a_tag_or_a_title_as_well(invoice):
    result = invoice.custom_xml_parts.fill({"/invoice[1]/customer[1]/contact[1]": "Ada Lovelace"})
    assert result.applied[0].xpath == "/invoice[1]/customer[1]/contact[1]"

    # the date control is the one with a w:tag in this document
    tagged = invoice.custom_xml_parts.fill({"/invoice[1]/invoicedate[1]": "2026-01-02T00:00:00Z"})
    assert tagged.applied[0].xpath == "/invoice[1]/invoicedate[1]"

    assert by_title(invoice, "/invoice[1]/customer[1]/contact[1]").text == "Ada Lovelace"


def test_fill_reports_a_key_that_matches_no_binding(invoice):
    result = invoice.custom_xml_parts.fill({"/invoice[1]/nothing[1]": "x"})

    assert result.applied == ()
    assert [entry.code for entry in result.skipped] == ["no_binding"]
    payload = result.to_dict()
    assert payload["skipped"][0]["key"] == "/invoice[1]/nothing[1]"
    assert json.loads(result.to_json())["bindings"]["operation"] == "apply_bindings"


def test_fill_with_a_string_replaces_the_whole_part(invoice):
    xml = invoice.custom_xml_parts.get_item(ITEM_ID).get_xml()
    replaced = xml.replace("<company>Contozo Inc</company>", "<company>Acme Ltd</company>")

    result = invoice.custom_xml_parts.fill(replaced)

    assert len(result.applied) == 1
    assert result.bindings.updated == 16
    assert by_title(invoice, "/invoice[1]/customer[1]/company[1]").text == "Acme Ltd"


def test_fill_with_a_string_refuses_when_it_cannot_tell_which_part():
    package = create_package()
    package.id_seed = SEED
    package.body.insert_paragraph("x")
    package.custom_xml_parts.add("<a><b>1</b></a>")
    package.custom_xml_parts.add("<c><d>2</d></c>")

    with pytest.raises(BindingError) as error:
        package.custom_xml_parts.fill("<a><b>3</b></a>")

    assert error.value.code == "binding.ambiguous_part"


def test_a_repeat_is_reported_but_not_expanded(invoice):
    """CR-003 section 17: a repeating section is a container, never bound."""
    before = len(
        next(
            c for c in invoice.body.content_controls if c.type == "RepeatingSection"
        ).repeating_section_content_control.items
    )

    invoice.custom_xml_parts.fill({"/invoice[1]/lines[1]/lineitem[1]/productcode[1]": "ITEM-9"})

    section = next(c for c in invoice.body.content_controls if c.type == "RepeatingSection")
    assert len(section.repeating_section_content_control.items) == before
    inner = by_title(invoice, "/invoice[1]/lines[1]/lineitem[1]/productcode[1]")
    assert inner.text == "ITEM-9"


# ---------------------------------------------------------------------------
# the trial, and determinism
# ---------------------------------------------------------------------------


def test_a_dry_run_of_a_fill_leaves_the_document_alone(invoice):
    invoice.body.paragraphs  # noqa: B018 - read the part first, as a caller would
    before = invoice.save()

    with invoice.dry_run() as trial:
        result = trial.custom_xml_parts.fill({"/invoice[1]/customer[1]/company[1]": "Acme Ltd"})
        assert result.bindings.updated == 16
        assert by_title(trial, "/invoice[1]/customer[1]/company[1]").text == "Acme Ltd"

    assert invoice.save() == before
    assert by_title(invoice, "/invoice[1]/customer[1]/company[1]").text == "Contozo Inc"
    assert "Acme Ltd" not in invoice.custom_xml_parts.get_item(ITEM_ID).get_xml()


def test_the_same_calls_make_the_same_bytes():
    """CR-003 section 3.4: a seeded package's ``add()`` is reproducible."""

    def build():
        package = create_package()
        package.id_seed = SEED
        package.body.insert_paragraph("Hello")
        part = package.custom_xml_parts.add('<g xmlns="urn:g"><to>World</to></g>')
        control = package.body.paragraphs[0].insert_content_control("PlainText")
        control.xml_mapping.set_mapping("/ns0:g[1]/ns0:to[1]", "xmlns:ns0='urn:g'", part)
        package.custom_xml_parts.apply_bindings()
        return package, part

    first, first_part = build()
    second, second_part = build()

    assert first_part.id == second_part.id
    assert first_part.id.startswith("{") and first_part.id == first_part.id.upper()
    assert first.save() == second.save()


def test_the_whole_workflow_end_to_end():
    """describe, fill, save, reload, read back: what a server does per request."""
    package = create_package()
    package.id_seed = SEED
    package.body.insert_paragraph("Dear ")
    part = package.custom_xml_parts.add(
        '<letter xmlns="urn:letter"><to>World</to><from>Ada</from></letter>'
    )
    mappings = "xmlns:ns0='urn:letter'"
    for name in ("to", "from"):
        control = package.body.paragraphs[0].get_range("End").insert_content_control("PlainText")
        assert control.xml_mapping.set_mapping(f"/ns0:letter[1]/ns0:{name}[1]", mappings, part)

    skeleton = package.custom_xml_parts.describe()
    assert [binding.xpath for binding in skeleton.bindings] == [
        "/ns0:letter[1]/ns0:to[1]",
        "/ns0:letter[1]/ns0:from[1]",
    ]
    assert [binding.value for binding in skeleton.bindings] == ["World", "Ada"]

    result = package.custom_xml_parts.fill(
        {"/ns0:letter[1]/ns0:to[1]": "Everyone", "/ns0:letter[1]/ns0:from[1]": "Grace"}
    )
    assert len(result.applied) == 2
    assert result.bindings.updated == 2

    saved = reloaded(package)
    assert [control.text for control in saved.body.content_controls] == ["Everyone", "Grace"]
    assert saved.custom_xml_parts.describe().bindings[0].value == "Everyone"


# ---------------------------------------------------------------------------
# the ChangeReport every mutation records (CR-003 decided question 4, §17.9)
# ---------------------------------------------------------------------------


def test_fill_records_exactly_one_report_for_the_whole_call(invoice):
    """`insert_markdown`'s rule: one call, one report, however much it writes."""
    invoice.changes.clear()

    result = invoice.custom_xml_parts.fill(
        {
            "/invoice[1]/customer[1]/company[1]": "Acme Ltd",
            "/invoice[1]/invoicenumber[1]": "INV-9",
            "/invoice[1]/nowhere[1]": "x",
        }
    )

    assert len(invoice.changes) == 1, [c.operation for c in invoice.changes]
    change = invoice.last_change
    assert change.operation == "fill"
    # the data part it wrote and the body part apply_bindings then wrote
    assert set(change.parts_touched) == {"/customXml/item1.xml", "/word/document.xml"}
    # the controls whose content changed
    assert len(change.addresses) >= result.bindings.updated
    assert any(address.startswith("body/") for address in change.addresses)
    # the counts, and the key that matched nothing, without a new field
    assert change.text_after == "2 set, 1 skipped, 16 bindings applied"
    assert any("/invoice[1]/nowhere[1]" in warning for warning in change.warnings)
    assert change.to_dict()["operation"] == "fill"


def test_a_dry_run_of_a_fill_reports_the_same_thing_over_the_trial(invoice):
    invoice.body.paragraphs  # noqa: B018 - read the part first, as a caller would
    invoice.changes.clear()

    with invoice.dry_run() as trial:
        trial.custom_xml_parts.fill({"/invoice[1]/customer[1]/company[1]": "Acme Ltd"})
        assert len(trial.changes) == 1
        change = trial.last_change
        assert change.operation == "fill"
        assert set(change.parts_touched) == {"/customXml/item1.xml", "/word/document.xml"}
        assert change.text_after == "1 set, 0 skipped, 16 bindings applied"

    # and the real package recorded nothing at all
    assert invoice.changes == []


def test_apply_bindings_and_the_reverse_each_record_one_report(invoice):
    invoice.custom_xml_parts.get_item(ITEM_ID).select_single_node(
        "/invoice/customer/company"
    ).text = "Acme Ltd"
    invoice.changes.clear()

    invoice.custom_xml_parts.apply_bindings()

    assert len(invoice.changes) == 1
    change = invoice.last_change
    assert change.operation == "apply_bindings"
    assert change.parts_touched == ("/word/document.xml",)
    assert len(change.addresses) == 16
    assert change.text_after == "16 of 20 bindings applied"
    assert len(change.warnings) == 4  # the picture, the rich text and the two containers

    invoice.changes.clear()
    invoice.custom_xml_parts.update_from_content_controls()

    assert len(invoice.changes) == 1
    back = invoice.last_change
    assert back.operation == "update_from_content_controls"
    # this direction writes the **data** parts, not the body
    assert "/word/document.xml" not in back.parts_touched


def test_the_node_and_part_mutations_each_record_one_report(invoice):
    part = invoice.custom_xml_parts.get_item(ITEM_ID)
    node = part.select_single_node("/invoice/customer/company")
    invoice.changes.clear()

    node.text = "Acme Ltd"

    assert len(invoice.changes) == 1
    change = invoice.last_change
    assert change.operation == "custom_xml_node.text"
    assert change.parts_touched == ("/customXml/item1.xml",)
    assert change.addresses == ("/invoice[1]/customer[1]/company[1]",)
    assert change.text_before == "Contozo Inc"
    assert change.text_after == "Acme Ltd"

    invoice.changes.clear()
    part.update_element("/invoice/customer/company", "<company>Zenith</company>")
    assert [c.operation for c in invoice.changes] == ["custom_xml_part.update_element"]
    assert invoice.last_change.parts_touched == ("/customXml/item1.xml",)

    invoice.changes.clear()
    part.insert_attribute("/invoice/customer", "kind", "trade")
    assert [c.operation for c in invoice.changes] == ["custom_xml_part.insert_attribute"]


def test_add_and_delete_name_the_parts_they_created(new_package):
    new_package.body.insert_paragraph("Hello")
    new_package.changes.clear()

    part = new_package.custom_xml_parts.add('<g xmlns="urn:g"><to>World</to></g>')

    assert len(new_package.changes) == 1
    change = new_package.last_change
    assert change.operation == "custom_xml_parts.add"
    assert set(change.parts_touched) == {
        "/customXml/item1.xml",
        "/customXml/itemProps1.xml",
        "/word/_rels/document.xml.rels",
    }
    assert change.addresses == (part.id,)

    control = new_package.body.paragraphs[0].insert_content_control("PlainText")
    control.xml_mapping.set_mapping("/ns0:g[1]/ns0:to[1]", "xmlns:ns0='urn:g'", part)
    new_package.changes.clear()

    part.delete()

    assert len(new_package.changes) == 1
    gone = new_package.last_change
    assert gone.operation == "custom_xml_part.delete"
    assert "/customXml/item1.xml" in gone.parts_touched
    assert "/customXml/itemProps1.xml" in gone.parts_touched
    assert any("/ns0:g[1]/ns0:to[1]" in warning for warning in gone.warnings)


def test_the_mapping_and_the_control_setters_each_record_one_report(new_package):
    part = new_package.custom_xml_parts.add('<g xmlns="urn:g"><to>World</to></g>')
    control = new_package.body.insert_paragraph("x").insert_content_control("CheckBox")
    new_package.changes.clear()

    assert control.xml_mapping.set_mapping("/ns0:g[1]/ns0:to[1]", "xmlns:ns0='urn:g'", part)
    assert [c.operation for c in new_package.changes] == ["xml_mapping.set_mapping"]
    assert new_package.last_change.addresses == (control.address,)
    assert new_package.last_change.parts_touched == ("/word/document.xml",)

    new_package.changes.clear()
    control.appearance = "Tags"
    control.color = "#FF0000"
    control.cannot_delete = True
    control.checkbox_content_control.is_checked = True
    assert [c.operation for c in new_package.changes] == [
        "control.appearance",
        "control.color",
        "control.cannot_delete",
        "checkbox.is_checked",
    ]
    assert all(change.addresses == (control.address,) for change in new_package.changes)

    new_package.changes.clear()
    control.xml_mapping.delete()
    assert [c.operation for c in new_package.changes] == ["xml_mapping.delete"]


def test_a_bound_insert_text_names_the_data_part_it_wrote(invoice):
    control = by_title(invoice, "/invoice[1]/invoicenumber[1]")
    invoice.changes.clear()

    control.insert_text("INV-9", location="Replace")

    assert len(invoice.changes) == 1
    change = invoice.last_change
    assert change.operation == "insert_text"
    # the write-through wrote the custom XML part, which re-marshals on save
    assert set(change.parts_touched) == {"/word/document.xml", "/customXml/item1.xml"}
    assert change.text_after == "INV-9"
