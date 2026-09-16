"""Budgets: a 200-page document read through a context window that cannot hold it.

CR-003 section 3.4 ("budgets everywhere") and section 7's token-budget test.
The sizes asserted here are the "stated size" section 7 asks for.
"""

from __future__ import annotations

import json

from conftest import big_document

from docx4j_py.model.content import TextExcerpt

#: The stated budgets, in bytes of UTF-8 JSON.
OUTLINE_BUDGET = 64 * 1024
HEADINGS_BUDGET = 8 * 1024
FIND_BUDGET = 8 * 1024


def test_a_200_page_document_outlines_inside_64_kb():
    package = big_document()
    outline = package.outline()

    assert outline.stats.paragraphs == 2180, "2,000 blocks, 20 of them three-row tables"
    assert outline.stats.tables == 20
    assert outline.stats.words > 30_000

    size = len(outline.to_json().encode("utf-8"))
    assert size < OUTLINE_BUDGET, f"{size} bytes"
    assert outline.truncated is True, "and it says the budget bit"
    assert len(outline.addresses()) == 300


def test_the_headings_view_is_inside_8_kb():
    package = big_document()
    outline = package.outline(headings_only=True)

    size = len(outline.to_json().encode("utf-8"))
    assert size < HEADINGS_BUDGET, f"{size} bytes"
    assert len(outline.entries) == 50, "a heading every fortieth paragraph"
    assert all(entry.level == 1 for entry in outline.entries)

    # the markdown is cheaper still, which is what to show a model
    assert len(outline.to_markdown().encode("utf-8")) < 2 * 1024


def test_find_at_its_default_limit_is_inside_8_kb():
    package = big_document()
    hits = package.find("lazy dog")

    assert len(hits) == 20, "the default limit"
    size = len(json.dumps([hit.to_dict() for hit in hits], separators=(",", ":")).encode("utf-8"))
    assert size < FIND_BUDGET, f"{size} bytes"


def test_limit_none_asks_for_the_whole_outline():
    package = big_document()
    outline = package.outline(limit=None)

    # 2,180 paragraphs (the cells' included) and 20 tables
    assert len(outline.addresses()) == 2180 + 20
    assert outline.truncated is True, "the texts are still cut at 80 characters"
    assert len(outline.to_json().encode("utf-8")) > OUTLINE_BUDGET


def test_the_budget_still_holds_when_every_paragraph_has_a_paraId():
    # a created document stamps every paragraph, which is 28 bytes an entry
    package = big_document()
    package.body.ensure_para_ids()

    outline = package.outline()
    size = len(outline.to_json().encode("utf-8"))
    assert size < OUTLINE_BUDGET, f"{size} bytes"
    headings = len(package.outline(headings_only=True).to_json().encode("utf-8"))
    assert headings < HEADINGS_BUDGET, f"{headings} bytes"


def test_get_text_takes_a_budget_and_text_budget_reports_it(report):
    body = report.body
    whole = body.get_text()

    assert body.get_text(max_chars=10) == whole[:10]

    excerpt = body.text_budget(10)
    assert isinstance(excerpt, TextExcerpt)
    assert excerpt.text == whole[:10]
    assert excerpt.chars == len(whole)
    assert excerpt.truncated is True
    assert str(excerpt) == excerpt.text
    assert excerpt.to_dict() == {"text": whole[:10], "chars": len(whole), "truncated": True}

    assert body.text_budget().truncated is False
    assert body.text_budget(len(whole)).truncated is False


def test_a_paragraph_has_the_same_two(report):
    paragraph = report.body.paragraphs[1]

    assert paragraph.get_text(max_chars=9) == "The quick"
    excerpt = paragraph.text_budget(9)
    assert (excerpt.text, excerpt.chars, excerpt.truncated) == ("The quick", 44, True)
    assert len(excerpt) == 9
