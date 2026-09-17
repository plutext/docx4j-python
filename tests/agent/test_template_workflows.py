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
