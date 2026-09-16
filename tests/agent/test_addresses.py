"""Addresses: the three forms, and what happens when one stops resolving.

CR-003 section 3.4, Phase D.
"""

from __future__ import annotations

import pytest
from conftest import WITH_PARA_IDS, WITHOUT_PARA_IDS, reloaded, sample

from docx4j_py.model.content import AddressError, Block, ContentError, Paragraph


def test_the_ordinal_is_one_index_per_block_step(report):
    body = report.body

    assert body.paragraphs[0].ordinal == "body/0"
    assert body.address_of(body[4]) == "body/4"
    assert body[4].name == "w:tbl"

    # a paragraph in a cell: table, row, cell, block --- four indices, and the
    # w:sdtContent levels of a control would be none at all
    cell_paragraph = report.paragraph_at("body/4/1/0/0")
    assert cell_paragraph.text == "North"
    assert cell_paragraph.ordinal == "body/4/1/0/0"


def test_para_id_first_and_the_ordinal_always(report):
    paragraph = report.body.paragraphs[0]

    assert paragraph.para_id, "a created document always stamps paragraph ids"
    assert paragraph.address == f"w14:{paragraph.para_id}"
    assert paragraph.ordinal == "body/0"
    assert report.paragraph_at(paragraph.address) == paragraph
    assert report.paragraph_at(para_id=paragraph.para_id) == paragraph


def test_a_loaded_document_without_para_ids_addresses_by_ordinal():
    package = sample(WITHOUT_PARA_IDS)
    paragraph = package.body.paragraphs[0]

    assert paragraph.para_id is None
    assert paragraph.address == paragraph.ordinal == "body/0"


def test_text_addressing_is_the_third_form():
    package = sample(WITH_PARA_IDS)
    text = package.body.paragraphs[0].text[:8]

    assert package.paragraph_at(contains=text).text.startswith(text)

    with pytest.raises(AddressError) as raised:
        package.paragraph_at(contains="no such words anywhere in this document")
    assert raised.value.code == "address.no_match"
    assert "find()" in raised.value.hint


def test_exactly_one_address_form_at_a_time(report):
    with pytest.raises(ContentError) as raised:
        report.paragraph_at("body/0", contains="Quarterly")
    assert raised.value.code == "address.ambiguous"

    with pytest.raises(ContentError):
        report.paragraph_at()


def test_an_address_that_is_gone_names_the_nearest_one(report):
    with pytest.raises(AddressError) as raised:
        report.element_at("body/99")
    assert raised.value.code == "address.not_found"
    assert "'body/6'" in str(raised.value), "the nearest surviving address is the last block"
    assert "outline()" in raised.value.hint

    # the nearest is chosen inside the block that shares the deepest path
    with pytest.raises(AddressError) as raised:
        report.element_at("body/4/1/0/9")
    assert "body/4/1/0/0" in str(raised.value)

    with pytest.raises(AddressError) as raised:
        report.element_at("w14:DEADBEEF")
    assert raised.value.code == "address.not_found"
    assert "para_id" in raised.value.hint


def test_a_paraId_address_survives_an_insert_and_an_ordinal_does_not(report):
    body = report.body
    target = body.paragraphs[2]
    address = target.address
    ordinal = target.ordinal
    assert ordinal == "body/2"

    body.insert_paragraph("Inserted at the top", location="Start")

    # the paraId still resolves to the same paragraph ...
    assert report.paragraph_at(address).text == target.text
    assert target.ordinal == "body/3"
    # ... and the ordinal now names its neighbour, which is why the report says
    # what moved (CR-003 section 3.4)
    assert report.paragraph_at(ordinal).text != target.text
    assert ("body/2", "body/3") in report.last_change.moved
    assert ("body/0", "body/1") in report.last_change.moved


def test_addresses_survive_a_save_and_a_reload(report):
    address = report.body.paragraphs[2].address
    text = report.body.paragraphs[2].text

    again = reloaded(report)
    assert again.paragraph_at(address).text == text


def test_element_at_reaches_a_header_through_the_package():
    package = sample("Headers.docx")
    prefixes = [body.prefix for body in package.bodies()]

    assert "body" in prefixes
    assert any(prefix.startswith("header:rId") for prefix in prefixes)
    header = next(p for p in prefixes if p.startswith("header:"))

    paragraph = package.paragraph_at(f"{header}/0")
    assert isinstance(paragraph, Paragraph)
    assert paragraph.ordinal == f"{header}/0"
    # and the body resolves it too, by handing it to its package
    assert package.body.paragraph_at(f"{header}/0") == paragraph


def test_element_at_refuses_a_table_as_a_paragraph(report):
    block = report.element_at("body/4")
    assert isinstance(block, Block)
    assert block.address == "body/4"
    assert repr(block) == "<Block body/4 w:tbl>"

    with pytest.raises(AddressError) as raised:
        report.paragraph_at("body/4")
    assert raised.value.code == "address.not_a_paragraph"
    assert "element_at" in raised.value.hint


def test_ensure_para_ids_stamps_a_legacy_document():
    package = sample(WITHOUT_PARA_IDS)
    body = package.body
    assert all(p.para_id is None for p in body.paragraphs)

    assigned = body.ensure_para_ids()
    assert len(assigned) == len(body.paragraphs)
    assert len(set(assigned)) == len(assigned), "no two paragraphs share an id"
    assert all(p.address.startswith("w14:") for p in body.paragraphs)

    # and it is deterministic: the same document and the same seed, twice
    twice = sample(WITHOUT_PARA_IDS)
    assert twice.body.ensure_para_ids() == assigned

    # the ids survive the round trip Word would make
    again = reloaded(package)
    assert [p.para_id for p in again.body.paragraphs] == assigned


def test_a_new_paragraph_joins_a_document_that_uses_para_ids():
    package = sample(WITH_PARA_IDS)
    taken = {p.para_id for p in package.body.paragraphs if p.para_id}
    assert taken

    made = package.body.insert_paragraph("added")
    assert made.para_id
    assert made.para_id not in taken
    assert package.last_change.created_para_ids == (made.para_id,)


def test_an_invalid_address_says_what_one_looks_like(report):
    with pytest.raises(ContentError) as raised:
        report.element_at("")
    assert raised.value.code == "address.invalid"
    assert "body/3" in str(raised.value)
