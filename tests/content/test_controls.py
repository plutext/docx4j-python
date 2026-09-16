"""``ContentControl``: the reads and ``delete``. CR-003 Phase C, section 3.2.

``samples/invoice2013.docx`` is the fixture section 7 names: it carries all
four ``w:sdt`` forms --- a block-level picture control, fifteen run-level plain
text controls, a date picker and a ``w15:repeatingSection`` whose ``w:tr`` is
wrapped in a row-level control --- with twenty bindings, three of them
``w15:dataBinding``. The typed kinds, the bindings and ``insert_content_control``
are Phase E; what is tested here is what a reader and a text editor need.
"""

from __future__ import annotations

import pytest
from conftest import reloaded, sample

from docx4j_py.model.content import ContentControl, InvalidTargetError, Range, Table


@pytest.fixture(scope="module")
def invoice():
    """``samples/invoice2013.docx``, loaded once for the reads."""
    return sample("invoice2013.docx")


def control_with(package, form: str) -> ContentControl:
    """The first control of a form."""
    return next(c for c in package.body.content_controls if c.form == form)


# ---------------------------------------------------------------------------
# reading
# ---------------------------------------------------------------------------


def test_every_control_is_found_in_document_order_nested_ones_included(invoice):
    controls = invoice.body.content_controls

    assert len(controls) == 22
    forms = {}
    for control in controls:
        forms[control.form] = forms.get(control.form, 0) + 1
    assert forms == {"block": 5, "run": 15, "row": 2}

    # document order: the addresses are non-decreasing down the body
    assert controls[0].address == "body/0"
    assert [control.address for control in controls] == sorted(
        [control.address for control in controls],
        key=lambda address: [int(part) for part in address.split("/")[1:]],
    )


def test_the_reads_against_known_values(invoice):
    control = next(
        c for c in invoice.body.content_controls if c.title.endswith("customer[1]/contact[1]")
    )

    assert control.form == "run"
    assert control.type == "PlainText"
    assert control.title == "/invoice[1]/customer[1]/contact[1]"
    assert control.tag == ""
    assert control.id == 2088573715
    assert control.text == "John Citizen"
    assert control.address == "body/2/0/0/0/0"
    assert control.name == "w:sdt"
    assert control.get_xml().startswith("<w:sdt ")

    assert control.to_dict() == {
        "kind": "control",
        "address": "body/2/0/0/0/0",
        "ordinal": "body/2/0/0/0/0",
        "type": "PlainText",
        "form": "run",
        "tag": "",
        "title": "/invoice[1]/customer[1]/contact[1]",
        "id": 2088573715,
        "text": "John Citizen",
    }


def test_an_untyped_control_is_rich_text(new_package):
    new_package.body.insert_xml(
        "<w:sdt><w:sdtPr><w:tag w:val='plain'/></w:sdtPr>"
        "<w:sdtContent><w:p><w:r><w:t>content</w:t></w:r></w:p></w:sdtContent></w:sdt>"
    )
    control = new_package.body.content_controls[0]

    assert control.type == "RichText", "w:sdtPr names no kind (CR-003 section 4)"
    assert control.form == "block"
    assert control.tag == "plain"
    assert control.title == ""
    assert control.id == 0


def test_the_four_forms_and_what_each_holds(invoice):
    block = control_with(invoice, "block")
    assert block.type == "Picture"
    assert block.paragraphs, "a block control answers with its own paragraphs"

    run = control_with(invoice, "run")
    assert len(run.paragraphs) == 1, "a run control answers with the paragraph it is in"
    assert run.paragraphs[0].text.startswith("John Citizen")

    row = control_with(invoice, "row")
    assert row.type == "RepeatingSection"
    tables = row.tables
    assert len(tables) == 1
    assert isinstance(tables[0], Table)
    assert tables[0].values[0][0] == "productcode"


def test_get_range_is_exact_for_a_run_control_and_the_first_paragraph_otherwise(invoice):
    run = control_with(invoice, "run")
    span = run.get_range()
    assert isinstance(span, Range)
    assert span.text == "John Citizen"
    assert (span.start, span.end) == (0, 12)
    assert run.get_range("Start").text == ""
    assert run.get_range("End").start == 12

    block = control_with(invoice, "block")
    assert block.get_range().paragraph == block.paragraphs[0]


def test_search_inside_a_control(invoice):
    run = next(c for c in invoice.body.content_controls if c.text == "John Citizen")
    assert [hit.text for hit in run.search("Citizen")] == ["Citizen"]
    assert run.search("Contozo") == []

    block = control_with(invoice, "block")
    assert isinstance(block.search("nothing here at all"), list)


def test_content_controls_of_a_paragraph_and_the_parent(invoice):
    paragraph = next(
        p for p in invoice.body.paragraphs if p.text.startswith("John Citizen")
    )
    controls = paragraph.content_controls

    assert len(controls) == 1
    assert controls[0].text == "John Citizen"
    assert controls[0].form == "run"

    inner = controls[0].paragraphs[0]
    assert inner == paragraph
    # the paragraph itself is not inside a run-level control, its runs are
    assert paragraph.parent_content_control is None

    # a paragraph inside a block-level control answers with it
    block = control_with(invoice, "block")
    assert block.paragraphs[0].parent_content_control == block


# ---------------------------------------------------------------------------
# editing
# ---------------------------------------------------------------------------


def test_insert_text_into_a_run_control(new_package):
    new_package.body.insert_xml(
        "<w:p><w:r><w:t>before </w:t></w:r>"
        "<w:sdt><w:sdtPr><w:tag w:val='name'/><w:text/></w:sdtPr>"
        "<w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent></w:sdt>"
        "<w:r><w:t> after</w:t></w:r></w:p>"
    )
    control = new_package.body.content_controls[0]
    assert control.form == "run"
    assert control.type == "PlainText"
    assert control.text == "Acme"

    control.insert_text(" Ltd", location="End")
    assert control.text == "Acme Ltd"
    control.insert_text("The ", location="Start")
    assert control.text == "The Acme Ltd"
    control.insert_text("Contozo", location="Replace")
    assert control.text == "Contozo"
    assert new_package.body.paragraphs[0].text == "before Contozo after"

    assert reloaded(new_package).body.content_controls[0].text == "Contozo"


def test_insert_into_a_block_control(new_package):
    new_package.body.insert_xml(
        "<w:sdt><w:sdtPr><w:tag w:val='notes'/></w:sdtPr>"
        "<w:sdtContent><w:p><w:r><w:t>one</w:t></w:r></w:p></w:sdtContent></w:sdt>"
    )
    control = new_package.body.content_controls[0]

    control.insert_paragraph("two")
    assert control.text == "one\ntwo"
    control.insert_paragraph("zero", location="Start")
    assert control.text == "zero\none\ntwo"
    control.insert_paragraph("beside", location="After")
    assert new_package.body.paragraphs[-1].text == "beside"
    assert control.text == "zero\none\ntwo"

    assert reloaded(new_package).body.content_controls[0].text == "zero\none\ntwo"


def test_a_row_control_refuses_what_it_cannot_hold(invoice):
    row = control_with(invoice, "row")

    with pytest.raises(InvalidTargetError) as raised:
        row.insert_paragraph("x", location="Start")
    assert raised.value.code == "control.wrong_form"
    assert "holds rows, not paragraphs" in raised.value.message
    assert "cell's body" in raised.value.hint

    with pytest.raises(InvalidTargetError) as raised:
        row.insert_text("x", location="Replace")
    assert raised.value.code == "control.wrong_form"
    assert "cells instead" in raised.value.hint


def test_delete_keeps_the_content_by_default(new_package):
    body = new_package.body
    body.insert_xml(
        "<w:sdt><w:sdtPr><w:tag w:val='notes'/></w:sdtPr><w:sdtContent>"
        "<w:p><w:r><w:t>kept</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>also kept</w:t></w:r></w:p>"
        "</w:sdtContent></w:sdt>"
    )
    control = body.content_controls[0]
    control.delete()

    assert body.content_controls == []
    assert [p.text for p in body.paragraphs] == ["kept", "also kept"]
    assert body.content[0].parent is body.container
    assert new_package.last_change.operation == "delete"

    back = reloaded(new_package)
    assert [p.text for p in back.body.paragraphs] == ["kept", "also kept"]


def test_delete_without_keeping_the_content(new_package):
    body = new_package.body
    body.insert_paragraph("outside")
    body.insert_xml(
        "<w:sdt><w:sdtPr><w:tag w:val='x'/></w:sdtPr>"
        "<w:sdtContent><w:p><w:r><w:t>gone</w:t></w:r></w:p></w:sdtContent></w:sdt>"
    )
    body.content_controls[0].delete(keep_content=False)

    assert [p.text for p in body.paragraphs] == ["outside"]
    assert reloaded(new_package).body.text == "outside"


def test_deleting_a_run_control_puts_its_runs_back_in_the_paragraph(new_package):
    new_package.body.insert_xml(
        "<w:p><w:r><w:t>before </w:t></w:r>"
        "<w:sdt><w:sdtPr><w:tag w:val='name'/></w:sdtPr>"
        "<w:sdtContent><w:r><w:t>inside</w:t></w:r></w:sdtContent></w:sdt>"
        "<w:r><w:t> after</w:t></w:r></w:p>"
    )
    control = new_package.body.content_controls[0]
    assert control.form == "run"
    control.delete()

    assert new_package.body.content_controls == []
    assert new_package.body.paragraphs[0].text == "before inside after"
    assert reloaded(new_package).body.text == "before inside after"


def test_views_compare_equal_hash_and_repr(invoice):
    one = control_with(invoice, "run")
    again = control_with(invoice, "run")

    assert one == again
    assert hash(one) == hash(again)
    assert len({one, again}) == 1
    assert repr(one) == (
        "<ContentControl body/2/0/0/0/0 PlainText "
        "'/invoice[1]/customer[1]/contact[1]' 'John Citizen'>"
    )
    assert str(one) == "John Citizen"


def test_view_for_hands_out_a_control(invoice):
    block = invoice.body[0]
    assert isinstance(block, ContentControl)
    assert invoice.element_at("body/0") == block
    assert block.address == "body/0"
