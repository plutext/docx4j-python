"""The typed content controls, the mapping and ``insert_content_control``.

CR-003 section 3.7 and section 4, Phase E. ``samples/invoice2013.docx`` carries
a checkbox, a date picker, a picture control and two ``w15:repeatingSection``\\ s
with twenty bindings; **no** document in ``samples/``, in docx4j or in
docx4j-core-ts carries a drop-down, a combo box or a group, so those three kinds
are exercised on documents this engine builds, as the TypeScript engine did.

Every mutation here is followed by save, reload and a read of the **markup**,
not only of the value through the view that wrote it (CR-003 section 16.11).
"""

from __future__ import annotations

import datetime

import pytest
from conftest import reloaded, sample

from docx4j_py.model.content import BindingError, ContentError, SpanError
from docx4j_py.model.customxml import (
    CheckboxContentControl,
    DatePickerContentControl,
    GroupContentControl,
    PictureContentControl,
    RepeatingSectionContentControl,
    format_date,
)

ITEM_ID = "{5D7BA57F-1E52-4637-9F82-2D4025768D4F}"

GREETING = '<greeting xmlns="http://example.com/g"><to>World</to></greeting>'
GREETING_NS = "xmlns:ns0='http://example.com/g'"
GREETING_PATH = "/ns0:greeting[1]/ns0:to[1]"


@pytest.fixture
def invoice():
    """``samples/invoice2013.docx``, a fresh copy per test: these tests mutate."""
    package = sample("invoice2013.docx")
    package.id_seed = 20260917
    return package


def by_title(package, title: str):
    """The control whose ``w:alias`` is `title`."""
    return next(control for control in package.body.content_controls if control.title == title)


def of_kind(package, kind: str):
    """Every control of a kind."""
    return [control for control in package.body.content_controls if control.type == kind]


# ---------------------------------------------------------------------------
# the kinds, on a document Word wrote
# ---------------------------------------------------------------------------


def test_the_kinds_the_invoice_carries(invoice):
    kinds = {control.type for control in invoice.body.content_controls}

    assert {"CheckBox", "DatePicker", "Picture", "RepeatingSection", "PlainText"} <= kinds

    checkbox = of_kind(invoice, "CheckBox")[0]
    view = checkbox.checkbox_content_control
    assert isinstance(view, CheckboxContentControl)
    assert view.is_checked is True
    assert view.checked_symbol == "☒"
    assert view.unchecked_symbol == "☐"
    assert view.font == "MS Gothic"
    # a kind view is None on a control of another kind
    assert checkbox.date_picker_content_control is None
    assert checkbox.group_content_control is None

    date = of_kind(invoice, "DatePicker")[0].date_picker_content_control
    assert isinstance(date, DatePickerContentControl)
    assert date.date_display_format == "d MMMM yyyy"
    assert date.date_display_locale == "en-AU"
    assert date.date_calendar_type == "gregorian"
    assert date.date_storage_format == "dateTime"
    assert date.full_date.year == 2015
    assert date.full_date.month == 1

    picture = of_kind(invoice, "Picture")[0].picture_content_control
    assert isinstance(picture, PictureContentControl)
    assert picture.inline_picture is not None

    sections = of_kind(invoice, "RepeatingSection")
    assert len(sections) == 2
    first = sections[0].repeating_section_content_control
    assert isinstance(first, RepeatingSectionContentControl)
    assert first.allow_insert_delete_section is True
    assert first.section_title == ""
    assert len(first.items) >= 1
    assert first.items[0].type == "RepeatingSectionItem"


def test_a_repeating_section_item_is_copied_as_words_plus_button_does(invoice):
    section = of_kind(invoice, "RepeatingSection")[0].repeating_section_content_control
    before = len(section.items)

    copy = section.insert_item_after(before - 1)

    assert copy.type == "RepeatingSectionItem"
    assert len(section.items) == before + 1
    # a fresh w:id, not the one it was copied from
    assert copy.id != section.items[0].id
    assert copy.id > 0

    saved = reloaded(invoice)
    again = of_kind(saved, "RepeatingSection")[0].repeating_section_content_control
    assert len(again.items) == before + 1

    with pytest.raises(BindingError) as error:
        section.insert_item_after(99)
    assert error.value.code == "binding.no_such_item"


def test_the_date_formatter_is_words_dateformat_pattern():
    when = datetime.datetime(2015, 1, 29, 14, 5, 9, tzinfo=datetime.UTC)

    assert format_date("d MMMM yyyy", when) == "29 January 2015"
    assert format_date("dd/MM/yyyy", when) == "29/01/2015"
    assert format_date("ddd d MMM yy", when) == "Thu 29 Jan 15"
    assert format_date("h:mm tt", when) == "2:05 PM"
    assert format_date("HH:mm:ss", when) == "14:05:09"
    assert format_date("'on' d MMMM", when) == "on 29 January"


# ---------------------------------------------------------------------------
# the w:sdtPr properties
# ---------------------------------------------------------------------------


def test_the_properties_round_trip_through_the_markup(new_package):
    control = new_package.body.insert_paragraph("Choose").insert_content_control("PlainText")

    control.appearance = "Tags"
    control.color = "#FF0000"
    control.cannot_delete = True
    control.cannot_edit = True
    control.remove_when_edited = True

    xml = control.get_xml()
    assert '<w15:appearance w15:val="tags"/>' in xml
    assert '<w15:color w:val="FF0000"/>' in xml
    assert '<w:lock w:val="sdtContentLocked"/>' in xml
    assert "<w:temporary/>" in xml

    # the two halves of one w:lock
    control.cannot_delete = False
    assert '<w:lock w:val="contentLocked"/>' in control.get_xml()
    control.cannot_edit = False
    assert "w:lock" not in control.get_xml()
    control.cannot_delete = True
    assert '<w:lock w:val="sdtLocked"/>' in control.get_xml()

    saved = reloaded(new_package).body.content_controls[0]
    assert saved.appearance == "Tags"
    assert saved.color == "#FF0000"
    assert saved.cannot_delete is True
    assert saved.cannot_edit is False
    assert saved.remove_when_edited is True


def test_an_unknown_appearance_is_refused(new_package):
    control = new_package.body.insert_paragraph("x").insert_content_control()

    with pytest.raises(BindingError) as error:
        control.appearance = "Sparkly"
    assert error.value.code == "binding.bad_appearance"
    assert control.appearance == "BoundingBox"
    assert control.color == ""


def test_placeholder_text_is_written_and_refused_over_content(new_package):
    body = new_package.body
    empty = body.insert_paragraph("").insert_content_control("PlainText")
    full = body.insert_paragraph("Ada Lovelace").insert_content_control("PlainText")

    assert empty.placeholder_text == ""
    empty.placeholder_text = "Your name"

    assert empty.is_showing_placeholder is True
    assert empty.placeholder_text == "Your name"
    xml = empty.get_xml()
    assert "<w:showingPlcHdr/>" in xml
    assert '<w:rStyle w:val="PlaceholderText"/>' in xml
    assert "Your name" in xml

    with pytest.raises(BindingError) as error:
        full.placeholder_text = "Your name"
    assert error.value.code == "binding.holds_content"
    assert "Ada Lovelace" in error.value.message
    assert full.text == "Ada Lovelace"

    saved = reloaded(new_package)
    again = saved.body.content_controls[0]
    assert again.is_showing_placeholder is True
    assert again.placeholder_text == "Your name"


# ---------------------------------------------------------------------------
# the lists and the group, on documents this engine builds
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("kind", ["DropDownList", "ComboBox"])
def test_a_list_control_holds_its_entries(new_package, kind):
    control = new_package.body.insert_paragraph("Choose").insert_content_control(kind)
    listing = (
        control.drop_down_list_content_control
        if kind == "DropDownList"
        else control.combo_box_content_control
    )

    assert listing.list_items == []
    listing.add_list_item("Apples", "apples")
    listing.add_list_item("Pears", "pears")
    listing.add_list_item("Plums", index=0)

    assert [(item.display_text, item.value) for item in listing.list_items] == [
        ("Plums", "Plums"),
        ("Apples", "apples"),
        ("Pears", "pears"),
    ]
    assert listing.list_items[1].index == 1
    assert '<w:listItem w:displayText="Apples" w:value="apples"/>' in control.get_xml()

    listing.list_items[0].delete()
    assert [item.display_text for item in listing.list_items] == ["Apples", "Pears"]

    saved = reloaded(new_package).body.content_controls[0]
    again = (
        saved.drop_down_list_content_control
        if kind == "DropDownList"
        else saved.combo_box_content_control
    )
    assert [item.value for item in again.list_items] == ["apples", "pears"]
    again.delete_all_list_items()
    assert again.list_items == []
    assert "w:listItem" not in saved.get_xml()


def test_a_group_control_wraps_the_whole_body(new_package):
    body = new_package.body
    body.insert_paragraph("The quick brown fox")
    body.insert_paragraph("Second")

    control = body.insert_content_control("Group")

    assert isinstance(control.group_content_control, GroupContentControl)
    assert len(body) == 1
    assert control.type == "Group"
    assert control.form == "block"
    assert control.text == "The quick brown fox\nSecond"
    assert "<w:group/>" in control.get_xml()

    saved = reloaded(new_package)
    assert len(saved.body) == 1
    assert saved.body.content_controls[0].type == "Group"
    assert saved.body.content_controls[0].text == "The quick brown fox\nSecond"


def test_a_checkbox_this_engine_builds_ticks_and_untick(new_package):
    control = new_package.body.insert_paragraph("").insert_content_control("CheckBox")
    view = control.checkbox_content_control

    assert view.is_checked is False
    view.is_checked = True

    assert control.text == "☒"
    xml = control.get_xml()
    assert '<w14:checked w14:val="1"/>' in xml
    assert '<w:rFonts w:hint="eastAsia" w:ascii="MS Gothic"' in xml

    saved = reloaded(new_package).body.content_controls[0]
    assert saved.checkbox_content_control.is_checked is True
    assert saved.text == "☒"
    saved.checkbox_content_control.is_checked = False
    assert saved.text == "☐"
    assert '<w14:checked w14:val="0"/>' in saved.get_xml()


# ---------------------------------------------------------------------------
# insert_content_control, at the three levels
# ---------------------------------------------------------------------------


def test_a_range_level_control_wraps_exactly_the_span(new_package):
    paragraph = new_package.body.insert_paragraph("The quick brown fox")

    control = paragraph.search("brown")[0].insert_content_control("PlainText")

    assert control.form == "run"
    assert control.type == "PlainText"
    assert control.text == "brown"
    assert paragraph.text == "The quick brown fox"
    assert control.id > 0
    assert "<w:text/>" in control.get_xml()

    saved = reloaded(new_package)
    again = saved.body.content_controls[0]
    assert again.form == "run"
    assert again.text == "brown"
    assert saved.body.paragraphs[0].text == "The quick brown fox"


def test_a_paragraph_level_control_wraps_it_in_place(new_package):
    body = new_package.body
    body.insert_paragraph("First")
    second = body.insert_paragraph("Second")

    control = second.insert_content_control("DatePicker")

    assert control.form == "block"
    assert control.text == "Second"
    assert len(body) == 2
    assert body.content[1] is control.element
    assert body[1] == control
    assert control.address == "body/1"

    saved = reloaded(new_package)
    from docx4j_py.traversal import element_name

    assert [element_name(block).rpartition("}")[2] for block in saved.body.content] == [
        "p",
        "sdt",
    ]
    assert saved.body.content_controls[0].type == "DatePicker"


def test_the_ids_are_unique_and_deterministic(new_package):
    body = new_package.body
    first = body.insert_paragraph("One").insert_content_control()
    second = body.insert_paragraph("Two").insert_content_control()

    assert first.id != second.id
    assert {first.id, second.id} == {1, 2}


def test_a_span_that_crosses_a_run_holder_is_refused(new_package):
    body = new_package.body
    body.insert_xml(
        '<w:p><w:r><w:t xml:space="preserve">see </w:t></w:r>'
        '<w:hyperlink r:id="rId99"><w:r><w:t>this page</w:t></w:r></w:hyperlink>'
        "<w:r><w:t> now</w:t></w:r></w:p>"
    )
    paragraph = body.paragraphs[-1]

    with pytest.raises(SpanError) as error:
        paragraph.get_range().insert_content_control("PlainText")

    assert error.value.code == "range.crosses_holder"
    assert "insert_content_control" in error.value.message


def test_a_repeating_section_is_refused_at_run_level(new_package):
    paragraph = new_package.body.insert_paragraph("The quick brown fox")

    with pytest.raises(BindingError) as error:
        paragraph.search("brown")[0].insert_content_control("RepeatingSection")

    assert error.value.code == "binding.form_mismatch"
    assert new_package.body.content_controls == []


def test_an_empty_body_cannot_be_wrapped(new_package):
    for block in list(new_package.body.content):
        new_package.body.content.remove(block)

    with pytest.raises(ContentError) as error:
        new_package.body.insert_content_control()

    assert error.value.code == "control.empty_body"


def test_an_inserted_control_is_not_a_revision(new_package):
    """CR-003 section 4: a content control is not a tracked change."""
    paragraph = new_package.body.insert_paragraph("The quick brown fox")
    new_package.change_tracking_mode = "TrackAll"

    control = paragraph.search("brown")[0].insert_content_control("PlainText")

    assert "<w:ins " not in control.get_xml()
    assert new_package.get_tracked_changes() == []
    # and deleting one is not a revision either (section 4, 16.6)
    control.delete()
    assert new_package.get_tracked_changes() == []
    assert new_package.body.paragraphs[0].text == "The quick brown fox"


# ---------------------------------------------------------------------------
# the XML mapping
# ---------------------------------------------------------------------------


def test_the_invoice_bindings_resolve_to_their_nodes(invoice):
    bound = [
        control for control in invoice.body.content_controls if control.xml_mapping.is_mapped
    ]

    assert len(bound) == 20
    assert {control.xml_mapping.store_item_id.lower() for control in bound} == {ITEM_ID.lower()}

    company = by_title(invoice, "/invoice[1]/customer[1]/company[1]")
    mapping = company.xml_mapping
    assert mapping.xpath == "/invoice[1]/customer[1]/company[1]"
    assert mapping.prefix_mappings == ""
    assert mapping.custom_xml_part.id == ITEM_ID
    assert mapping.custom_xml_node.text == "Contozo Inc"
    assert mapping.to_dict()["resolves"] is True

    # three of the twenty are w15:dataBinding, and they are read all the same
    w15 = [
        control
        for control in bound
        if "w15:dataBinding" in control.get_xml().split("<w:sdtContent>")[0]
    ]
    assert len(w15) == 3


def test_set_mapping_resolves_before_it_writes(new_package):
    control = new_package.body.insert_paragraph("x").insert_content_control("PlainText")
    part = new_package.custom_xml_parts.add(GREETING)

    assert control.xml_mapping.is_mapped is False
    assert control.xml_mapping.set_mapping(GREETING_PATH, GREETING_NS, part) is True
    assert control.xml_mapping.xpath == GREETING_PATH
    assert control.xml_mapping.store_item_id == part.id
    assert control.xml_mapping.custom_xml_node.text == "World"
    assert f'w:xpath="{GREETING_PATH}"' in control.get_xml()

    # an XPath that selects nothing writes nothing and says so
    assert control.xml_mapping.set_mapping("/ns0:greeting[1]/ns0:nope[1]", GREETING_NS) is False
    assert control.xml_mapping.xpath == GREETING_PATH

    saved = reloaded(new_package).body.content_controls[0]
    assert saved.xml_mapping.xpath == GREETING_PATH
    assert saved.xml_mapping.custom_xml_node.text == "World"


def test_set_mapping_by_node_uses_the_nodes_own_path(new_package):
    control = new_package.body.insert_paragraph("x").insert_content_control("PlainText")
    part = new_package.custom_xml_parts.add(GREETING)
    node = part.select_single_node(GREETING_PATH)

    assert control.xml_mapping.set_mapping_by_node(node) is True

    assert control.xml_mapping.xpath == GREETING_PATH
    assert control.xml_mapping.prefix_mappings == "xmlns:ns0='http://example.com/g'"
    assert control.xml_mapping.custom_xml_node == node


def test_delete_removes_the_binding_and_keeps_the_content(invoice):
    company = by_title(invoice, "/invoice[1]/customer[1]/company[1]")
    text = company.text

    company.xml_mapping.delete()

    assert company.xml_mapping.is_mapped is False
    assert company.text == text
    assert "w:dataBinding" not in company.get_xml()
    saved = reloaded(invoice)
    again = by_title(saved, "/invoice[1]/customer[1]/company[1]")
    assert again.xml_mapping.is_mapped is False
    assert again.text == text


def test_a_w15_binding_keeps_its_namespace_when_it_is_rewritten(invoice):
    section = of_kind(invoice, "RepeatingSection")[0]
    assert "w15:dataBinding" in section.get_xml().split("<w:sdtContent>")[0]

    assert section.xml_mapping.set_mapping("/invoice[1]/lines[1]/lineitem[1]", "") is True

    properties = section.get_xml().split("<w:sdtContent>")[0]
    assert "w15:dataBinding" in properties
    # exactly one binding element, not a w: one added beside the w15: one
    assert properties.count("dataBinding") == 1


# ---------------------------------------------------------------------------
# applying the bindings, both ways
# ---------------------------------------------------------------------------


def test_apply_bindings_pushes_the_data_into_the_controls(invoice):
    part = invoice.custom_xml_parts.get_item(ITEM_ID)
    part.select_single_node("/invoice/customer/company").text = "Acme Ltd"

    result = invoice.custom_xml_parts.apply_bindings()

    assert result.bound == 20
    assert result.updated == 16
    assert {entry.code for entry in result.skipped} == {"picture", "container", "rich_text"}
    assert by_title(invoice, "/invoice[1]/customer[1]/company[1]").text == "Acme Ltd"
    # the date is formatted, the checkbox shows its glyph
    assert of_kind(invoice, "DatePicker")[0].text == "29 January 2015"
    assert of_kind(invoice, "CheckBox")[0].text == "☒"
    # and the containers still hold their items
    assert len(of_kind(invoice, "RepeatingSection")[0].repeating_section_content_control.items) >= 1

    saved = reloaded(invoice)
    assert by_title(saved, "/invoice[1]/customer[1]/company[1]").text == "Acme Ltd"
    assert "Acme Ltd" in saved.custom_xml_parts.get_item(ITEM_ID).get_xml()


def test_update_from_content_controls_writes_the_controls_back(invoice):
    checkbox = of_kind(invoice, "CheckBox")[0]
    checkbox.checkbox_content_control.is_checked = False
    date = of_kind(invoice, "DatePicker")[0]

    result = invoice.custom_xml_parts.update_from_content_controls()

    assert result.bound == 20
    assert result.updated >= 1
    # the word, not the glyph, and the stored form, not the display
    assert checkbox.xml_mapping.custom_xml_node.text == "false"
    assert date.xml_mapping.custom_xml_node.text == "2015-01-29T00:00:00"
    assert "<applies" not in invoice.custom_xml_parts.get_item(ITEM_ID).get_xml()

    saved = reloaded(invoice)
    assert saved.custom_xml_parts.get_item(ITEM_ID).select_single_node(
        "/invoice/VAT/@applies"
    ).text == "false"


def test_a_list_shows_the_display_text_and_stores_the_value(new_package):
    control = new_package.body.insert_paragraph("x").insert_content_control("DropDownList")
    listing = control.drop_down_list_content_control
    listing.add_list_item("Apples", "apples")
    listing.add_list_item("Pears", "pears")
    part = new_package.custom_xml_parts.add(GREETING)
    control.xml_mapping.set_mapping(GREETING_PATH, GREETING_NS, part)
    part.select_single_node(GREETING_PATH).text = "pears"

    new_package.custom_xml_parts.apply_bindings()

    assert control.text == "Pears"
    assert listing.last_value == "pears"

    listing.list_items[0].display_text = "Pears"  # make the display ambiguous in the other order
    control.set_bound_content([])
    control.insert_text("Apples")
    new_package.custom_xml_parts.update_from_content_controls()
    assert part.select_single_node(GREETING_PATH).text == "Apples"


def test_a_bound_controls_insert_text_writes_through(invoice):
    control = by_title(invoice, "/invoice[1]/invoicenumber[1]")

    control.insert_text("INV-9", location="Replace")

    assert control.text == "INV-9"
    assert control.xml_mapping.custom_xml_node.text == "INV-9"
    assert "INV-9" in invoice.custom_xml_parts.get_item(ITEM_ID).get_xml()

    saved = reloaded(invoice)
    again = by_title(saved, "/invoice[1]/invoicenumber[1]")
    assert again.text == "INV-9"
    assert again.xml_mapping.custom_xml_node.text == "INV-9"


def test_the_write_through_happens_with_tracking_on_too(invoice):
    """CR-003 section 16.6: the hook sits outside the tracked branch."""
    invoice.change_tracking_mode = "TrackAll"
    control = by_title(invoice, "/invoice[1]/invoicenumber[1]")

    control.insert_text("INV-10", location="Replace")

    # the data took the new value whatever the mode is
    assert control.xml_mapping.custom_xml_node.text == "INV-10"
    # and the document carries the revision markup as well
    assert "<w:ins " in control.get_xml()
    assert [change.type for change in invoice.get_tracked_changes()]

    saved = reloaded(invoice)
    assert saved.custom_xml_parts.get_item(ITEM_ID).select_single_node(
        "/invoice/invoicenumber"
    ).text == "INV-10"


def test_a_binding_whose_part_is_gone_is_reported_not_raised(new_package):
    control = new_package.body.insert_paragraph("x").insert_content_control("PlainText")
    part = new_package.custom_xml_parts.add(GREETING)
    control.xml_mapping.set_mapping(GREETING_PATH, GREETING_NS, part)
    part.delete()
    # put the binding back by hand: delete() unlinked it
    part2 = new_package.custom_xml_parts.add(GREETING)
    control.xml_mapping.set_mapping(GREETING_PATH, GREETING_NS, part2)
    control.xml_mapping.data_binding.store_item_id = "{00000000-0000-0000-0000-000000000000}"

    result = new_package.custom_xml_parts.apply_bindings()

    assert result.updated == 0
    assert [entry.code for entry in result.skipped] == ["no_part"]
    assert "00000000" in result.skipped[0].reason


def test_an_xpath_that_no_longer_resolves_is_reported(new_package):
    control = new_package.body.insert_paragraph("x").insert_content_control("PlainText")
    part = new_package.custom_xml_parts.add(GREETING)
    control.xml_mapping.set_mapping(GREETING_PATH, GREETING_NS, part)
    part.set_xml('<greeting xmlns="http://example.com/g"><other>x</other></greeting>')

    result = new_package.custom_xml_parts.apply_bindings()

    assert [entry.code for entry in result.skipped] == ["no_match"]
    assert result.to_dict()["bound"] == 1


def test_the_result_is_a_frozen_dataclass_with_a_json_view(invoice):
    result = invoice.custom_xml_parts.apply_bindings()

    payload = result.to_dict()
    assert payload["operation"] == "apply_bindings"
    assert payload["bound"] == 20
    assert set(payload) == {"operation", "bound", "updated", "applied", "skipped"}
    assert result.to_json().startswith("{")
    with pytest.raises(Exception):  # noqa: B017 - frozen
        result.operation = "no"
