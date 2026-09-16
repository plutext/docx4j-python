"""Markdown in and out: the constructs, one at a time. CR-003 section 3.5.

Reads against known values for each construct, every insert followed by save,
reload and read back, and the promise CR-003 section 7 makes for every phase:
an untouched part is still byte-identical through the content API, and the one
part the markdown *does* touch --- ``styles.xml``, when a style has to be added
--- is reported rather than touched quietly.

The oracle is Java docx4j's ``docx4j-markdown``; where this departs from it,
CR-003 section 13 says why.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest
from conftest import SAMPLES, part_bytes, reloaded, sample

from docx4j_py import WordprocessingMLPackage, load
from docx4j_py.model.content import ContentError, TextExcerpt
from docx4j_py.model.markdown import (
    ADDRESS_COMMENT,
    STYLE_IDS,
    MarkdownView,
    address_comment,
    ensure_style,
    style_ids_of,
    table_markdown,
)

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

#: The documents copied for this phase (see ``tests/README.md``).
FIXTURES = ROOT / "tests" / "fixtures"


def fixture(name: str) -> WordprocessingMLPackage:
    """Load one of ``tests/fixtures/``."""
    return load(FIXTURES / f"{name}.docx")


@pytest.fixture
def report(new_package: WordprocessingMLPackage) -> WordprocessingMLPackage:
    """A small document with headings, body text and a table: enough to address."""
    body = new_package.body
    body.insert_paragraph("Quarterly Report", style="Heading 1")
    body.insert_paragraph("The quick brown fox jumps over the lazy dog.")
    body.insert_paragraph("Revenue", style="Heading 2")
    body.insert_xml(
        "<w:tbl><w:tblPr/><w:tblGrid/>"
        "<w:tr><w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>Total</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl>"
    )
    new_package.changes.clear()
    return new_package


# ---------------------------------------------------------------------------
# out: one construct at a time
# ---------------------------------------------------------------------------


def test_a_heading_is_hashes_from_the_shared_heading_level(body):
    body.insert_paragraph("Title", style="Heading 1")
    body.insert_paragraph("Sub", style="Heading 2")
    body.insert_paragraph("Plain")

    assert body.to_markdown() == "# Title\n\n## Sub\n\nPlain"


def test_an_outline_level_is_a_heading_without_a_heading_style(body):
    body.insert_xml(
        '<w:p><w:pPr><w:outlineLvl w:val="2"/></w:pPr>'
        "<w:r><w:t>By outline level</w:t></w:r></w:p>"
    )

    assert body.to_markdown() == "### By outline level"


def test_bold_italic_strike_and_code(body):
    body.insert_xml(
        "<w:p>"
        "<w:r><w:t>plain </w:t></w:r>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>"
        "<w:r><w:t> </w:t></w:r>"
        "<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>"
        "<w:r><w:t> </w:t></w:r>"
        "<w:r><w:rPr><w:strike/></w:rPr><w:t>gone</w:t></w:r>"
        "<w:r><w:t> </w:t></w:r>"
        '<w:r><w:rPr><w:rStyle w:val="CodeChar"/></w:rPr><w:t>x = 1</w:t></w:r>'
        "</w:p>"
    )

    assert body.to_markdown() == "plain **bold** *italic* ~~gone~~ `x = 1`"


def test_a_monospace_font_is_code_when_no_style_says_so(body):
    body.insert_xml(
        "<w:p><w:r><w:rPr><w:rFonts w:ascii=\"Consolas\" w:hAnsi=\"Consolas\"/></w:rPr>"
        "<w:t>mono</w:t></w:r></w:p>"
    )

    assert body.to_markdown() == "`mono`"


def test_two_bold_runs_in_a_row_are_one_pair_of_markers(body):
    body.insert_xml(
        "<w:p>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>one </w:t></w:r>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>two</w:t></w:r>"
        "</w:p>"
    )

    assert body.to_markdown() == "**one two**"


def test_a_hyperlink_reads_its_relationship_and_an_anchor_reads_its_anchor():
    package = fixture("hyperlink")

    assert package.to_markdown() == (
        "This document contains a [hyperlink](http://slashdot.org/)"
    )


def test_an_internal_link_is_a_fragment():
    package = sample("toc.docx")

    assert "[Article 1\t92](#_Toc467580795)" in package.to_markdown()


def test_a_list_is_bullets_or_numbers_from_the_numbering_parts_numFmt():
    package = fixture("lists")
    markdown = package.to_markdown()

    # numId 1 is w:numFmt="decimal" at ilvl 0, so the items are numbered; the
    # empty paragraphs between them close the list each time, as Java's
    # WmlToMarkdown does, so each restarts at w:start
    assert len(re.findall(r"^1\. ", markdown, re.MULTILINE)) == 4
    assert "- " not in markdown, "numId 1 is decimal, not a bullet"


def test_a_bullet_list_nests_by_ilvl(new_package):
    body = new_package.body
    body.insert_markdown("- one\n- two\n  - nested\n")

    assert body.to_markdown() == "- one\n- two\n  - nested"


def test_a_table_is_a_gfm_pipe_table(new_package):
    body = new_package.body
    body.insert_xml(
        "<w:tbl><w:tblPr/><w:tblGrid/>"
        "<w:tr><w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>Total</w:t></w:r></w:p></w:tc></w:tr>"
        "<w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>120</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl>"
    )

    assert body.to_markdown() == (
        "| Region | Total |\n| --- | --- |\n| North | 120 |"
    )


def test_the_table_renderer_works_over_the_element_alone_for_phase_c(new_package):
    body = new_package.body
    blocks = body.insert_xml(
        "<w:tbl><w:tblPr/><w:tblGrid/>"
        "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl>"
    )

    # Phase C's Table.to_markdown hangs on this, so it takes the element and
    # any context (a Body, or none at all)
    assert table_markdown(blocks[0].element, body) == "| A |\n| --- |"
    assert table_markdown(blocks[0].element) == "| A |\n| --- |"


def test_a_header_rows_bold_is_convention_and_is_not_marked(new_package):
    body = new_package.body
    body.insert_markdown("| A | B |\n| --- | --- |\n| 1 | 2 |\n")

    # the import bolds the header row (TableGrid has no header formatting of
    # its own), and the export must not read that back as markdown emphasis
    assert body.to_markdown() == "| A | B |\n| --- | --- |\n| 1 | 2 |"


def test_an_image_is_its_part_name_relative_to_word():
    package = sample("Images.docx")

    assert package.to_markdown() == (
        "![Picture 1](media/image1.png)\n\n![pangolin.jpeg](media/image2.jpeg)"
    )


def test_a_footnote_reference_is_a_gfm_footnote_with_its_definition_at_the_end():
    package = fixture("footnotes")
    markdown = package.to_markdown()

    assert markdown.startswith("Hello hello[^1]")
    assert markdown.endswith("[^1]: Hello")


def test_a_paragraph_renders_on_its_own_without_the_footnote_definitions():
    package = fixture("footnotes")
    first = package.body.paragraphs[0]

    assert first.to_markdown() == "Hello hello[^1]"


def test_the_accepted_view_drops_deletions_and_the_markup_view_is_criticmarkup():
    accepted = sample("sample-docx.docx").to_markdown()
    markup = sample("sample-docx.docx").to_markdown(view="markup")

    assert "Here is some change tracking. An insertion Followed by." in accepted
    assert (
        "Here is some change tracking. {++An insertion++} Followed by{-- A deletion--}."
        in markup
    )


def test_a_comment_is_criticmarkup_in_the_markup_view_only():
    assert fixture("comments").to_markdown() == "One paragraph"
    assert fixture("comments").to_markdown(view="markup") == "One paragraph{>>One comment<<}"


def test_the_view_values_are_a_literal_with_a_strenum_beside_them(body):
    body.insert_paragraph("x")

    assert MarkdownView.ACCEPTED == "accepted"
    assert body.to_markdown(view=MarkdownView.MARKUP) == "x"
    with pytest.raises(ContentError) as error:
        body.to_markdown(view="original")
    assert error.value.code == "markdown.view_invalid"


def test_a_code_style_paragraph_becomes_a_fence_and_consecutive_ones_merge(body):
    for line in ("x = 1", "y = 2"):
        body.insert_paragraph(line).style_id = "SourceCode"

    assert body.to_markdown() == "```\nx = 1\ny = 2\n```"


def test_a_quote_style_paragraph_is_a_block_quote(body):
    body.insert_paragraph("Quoted.").style_id = "Quote"

    assert body.to_markdown() == "> Quoted."


def test_markup_that_would_be_read_back_as_markup_is_escaped(body):
    body.insert_paragraph("a * b _ c [d] `e` \\f")
    body.insert_paragraph("1. not a list")
    body.insert_paragraph("- not a bullet")

    assert body.to_markdown() == (
        "a \\* b \\_ c \\[d\\] \\`e\\` \\\\f\n\n\\1. not a list\n\n\\- not a bullet"
    )


def test_an_empty_paragraph_contributes_nothing(body):
    body.insert_paragraph("one")
    body.insert_paragraph("")
    body.insert_paragraph("two")

    assert body.to_markdown() == "one\n\ntwo"


# ---------------------------------------------------------------------------
# addresses
# ---------------------------------------------------------------------------


def test_each_block_carries_its_address_on_its_own_line_before_it(report):
    markdown = report.body.to_markdown(addresses=True)

    lines = markdown.splitlines()
    assert lines[0] == address_comment(report.body.paragraphs[0].address)
    assert lines[1] == "# Quarterly Report"


def test_the_documented_regex_recovers_every_address_and_element_at_takes_it(report):
    markdown = report.body.to_markdown(addresses=True)

    found = [match.group("address") for match in ADDRESS_COMMENT.finditer(markdown)]
    assert found, "the export emitted no address comments"
    for address in found:
        report.element_at(address)  # raises AddressError if it does not resolve
    assert found == [block.address for block in report.body.iter_blocks()]


def test_the_address_is_the_paraid_and_not_the_ordinal(report):
    markdown = report.body.to_markdown(addresses=True)

    first = ADDRESS_COMMENT.search(markdown).group("address")
    assert first.startswith("w14:"), "CR-003 section 12.7: emit address, not ordinal"
    assert first == report.body.paragraphs[0].address


def test_a_legacy_document_reports_ordinals_instead():
    package = sample("2010-sample1.docx")
    markdown = package.to_markdown(addresses=True)

    assert ADDRESS_COMMENT.search(markdown).group("address") == "body/0"


def test_an_address_comment_inside_a_list_is_indented_with_the_item(new_package):
    new_package.body.insert_markdown("- one\n  - nested\n")

    markdown = new_package.body.to_markdown(addresses=True)
    assert "\n  <!-- " in markdown, "the nested item's comment is indented as it is"


# ---------------------------------------------------------------------------
# budgets
# ---------------------------------------------------------------------------


def test_max_chars_cuts_at_a_block_boundary(report):
    whole = report.body.to_markdown()
    cut = report.body.to_markdown(max_chars=40)

    assert whole.startswith(cut)
    assert len(cut) <= 40
    assert cut in ("# Quarterly Report", whole[: len(cut)])
    assert not cut.endswith(" ")


def test_markdown_budget_says_whether_the_budget_bit(report):
    whole = report.markdown_budget()
    cut = report.markdown_budget(40)

    assert isinstance(whole, TextExcerpt)
    assert whole.truncated is False
    assert whole.chars == len(report.to_markdown())
    assert cut.truncated is True
    assert cut.chars == whole.chars


def test_a_budget_smaller_than_the_first_block_truncates_inside_it(report):
    cut = report.markdown_budget(5)

    assert cut.text == "# Qua"
    assert cut.truncated is True


# ---------------------------------------------------------------------------
# in
# ---------------------------------------------------------------------------


def test_insert_markdown_returns_the_inserted_views_in_document_order(body):
    views = body.insert_markdown("# One\n\nTwo\n\n- three\n")

    assert [view.text for view in views] == ["One", "Two", "three"]
    assert [view.address for view in views] == [block.address for block in body.iter_blocks()]


def test_a_heading_gets_the_documents_own_heading_style(body):
    views = body.insert_markdown("## Sub\n")

    assert views[0].style_id == "Heading2"
    assert views[0].style == "Heading 2"


def test_emphasis_becomes_run_properties(body):
    body.insert_markdown("Plain **bold** *italic* ~~gone~~ `code`.\n")

    paragraph = body.paragraphs[0]
    assert paragraph.text == "Plain bold italic gone code."
    runs = paragraph.runs
    assert [run.r_pr is not None and run.r_pr.b is not None for run in runs][1] is True
    assert [run.r_pr is not None and run.r_pr.i is not None for run in runs][3] is True
    assert [run.r_pr is not None and run.r_pr.strike is not None for run in runs][5] is True
    assert runs[7].r_pr.r_style.val == "CodeChar"


def test_a_link_becomes_a_hyperlink_with_an_external_relationship(body, new_package):
    body.insert_markdown("See [docx4j](https://www.docx4java.org/).\n")

    relationships = new_package.main_document_part.relationships_part
    targets = [rel.target for rel in relationships.list]
    assert "https://www.docx4java.org/" in targets
    assert body.to_markdown() == "See [docx4j](https://www.docx4java.org/)."


def test_a_list_gets_a_numPr_and_the_numbering_part_is_created(body, new_package):
    assert new_package.numbering_definitions_part is None

    body.insert_markdown("- one\n- two\n")

    part = new_package.numbering_definitions_part
    assert part is not None
    assert len(part.contents.abstract_num) == 1
    assert len(part.contents.lvl if False else part.contents.abstract_num[0].lvl) == 9
    num_pr = body.paragraphs[0].element.p_pr.num_pr
    assert int(num_pr.num_id.val) == int(part.contents.num[0].num_id)
    assert int(num_pr.ilvl.val) == 0


def test_an_ordered_list_keeps_its_start_value(body, new_package):
    body.insert_markdown("3. three\n4. four\n")

    part = new_package.numbering_definitions_part
    assert int(part.contents.abstract_num[0].lvl[0].start.val) == 3
    assert body.to_markdown() == "3. three\n4. four"


def test_a_nested_list_of_the_other_kind_takes_the_signature_javas_scan_gives(
    body, new_package
):
    body.insert_markdown("- one\n  1. a\n  2. b\n")

    formats = [str(level.num_fmt.val) for level in
               new_package.numbering_definitions_part.contents.abstract_num[0].lvl[:2]]
    assert formats == ["bullet", "decimal"]


def test_a_table_becomes_a_tbl_with_a_grid_sized_from_the_sectPr(body, new_package):
    body.insert_markdown("| A | B |\n| :-- | --: |\n| 1 | 2 |\n")

    table = body.content[0]
    widths = [int(col.w) for col in table.tbl_grid.grid_col]
    assert len(widths) == 2
    assert sum(widths) == 11907 - 1440 - 1440, "A4 minus the created margins"
    assert str(table.content[1].content[1].content[0].p_pr.jc.val) == "right"


def test_a_code_block_becomes_source_code_paragraphs(body):
    body.insert_markdown("```python\nx = 1\ny = 2\n```\n")

    assert body.paragraphs[0].style_id == "SourceCode"
    assert body.to_markdown() == "```\nx = 1\ny = 2\n```"


def test_a_block_quote_becomes_a_quote_styled_paragraph(body):
    body.insert_markdown("> quoted\n")

    assert body.paragraphs[0].style_id == "Quote"


def test_a_thematic_break_is_an_empty_paragraph_with_a_bottom_border(body):
    body.insert_markdown("---\n")

    assert body.content[0].p_pr.p_bdr.bottom is not None


def test_an_image_is_never_fetched_and_is_reported(body, new_package):
    body.insert_markdown("![a pangolin](https://example.com/pangolin.png)\n")

    assert body.to_markdown() == "[a pangolin](https://example.com/pangolin.png)"
    assert any("not fetched" in w for w in new_package.last_change.warnings)


def test_an_html_block_is_skipped_and_reported(body, new_package):
    body.insert_markdown("Text.\n\n<div>raw</div>\n")

    assert body.to_markdown() == "Text."
    assert any("HTML block was skipped" in w for w in new_package.last_change.warnings)


def test_markdown_that_makes_no_block_is_reported_rather_than_dropped(body, new_package):
    assert body.insert_markdown("   \n") == []
    assert new_package.last_change.warnings == ("this markdown produced no block-level content",)


def test_gfm_footnote_import_is_out_of_scope_and_stays_literal(body):
    # markdown-it's core has no footnote rule and mdit-py-plugins is not a
    # dependency this CR takes (CR-003 section 13): the marker stays text
    body.insert_markdown("A claim[^1]\n\n[^1]: the evidence\n")

    assert body.paragraphs[0].text == "A claim[^1]"


def test_one_recording_covers_the_whole_fragment(new_package):
    new_package.changes.clear()
    new_package.body.insert_markdown("# One\n\nTwo\n\nThree\n\n- four\n")

    assert len(new_package.changes) == 1
    report = new_package.last_change
    assert report.operation == "insert_markdown"
    assert len(report.addresses) == 4


def test_location_start_puts_the_fragment_first(body):
    body.insert_paragraph("Existing")
    body.insert_markdown("# First\n", location="Start")

    assert [p.text for p in body.paragraphs] == ["First", "Existing"]


def test_a_paragraph_inserts_after_itself_by_default(body):
    first = body.insert_paragraph("One")
    body.insert_paragraph("Three")
    views = first.insert_markdown("Two\n")

    assert [p.text for p in body.paragraphs] == ["One", "Two", "Three"]
    assert views[0].text == "Two"


def test_a_single_paragraph_fragment_merges_into_the_paragraph_at_end(body):
    paragraph = body.insert_paragraph("Start. ")
    views = paragraph.insert_markdown("And **more**.", location="End")

    assert views == [paragraph]
    assert paragraph.text == "Start. And more."
    assert len(body.paragraphs) == 1


# ---------------------------------------------------------------------------
# what is touched, and what is not
# ---------------------------------------------------------------------------


def test_reading_a_document_as_markdown_leaves_every_other_part_byte_for_byte():
    untouched = sample("tables.docx")
    before = part_bytes(untouched.save())

    package = sample("tables.docx")
    package.to_markdown()
    after = part_bytes(package.save())

    changed = [name for name, data in before.items() if after.get(name) != data]
    assert changed == ["word/document.xml"], "only the part the body needed"


def test_inserting_markdown_that_needs_no_new_style_touches_only_the_body():
    package = sample("2010-sample1.docx")
    package.id_seed = 20260917
    package.body  # the part the edit is in
    package.style_definitions_part.contents  # already unmarshalled either way
    before = part_bytes(package.save())

    package.body.insert_markdown("Just a plain paragraph.\n")
    after = part_bytes(package.save())

    changed = [name for name, data in before.items() if after.get(name) != data]
    assert changed == ["word/document.xml"]
    assert package.last_change.parts_touched == ("/word/document.xml",)


def test_a_styles_part_touch_is_reported(new_package):
    body = new_package.body
    assert "Quote" not in style_ids_of(new_package.style_definitions_part)

    body.insert_markdown("> quoted\n")

    assert "Quote" in style_ids_of(new_package.style_definitions_part)
    assert "/word/styles.xml" in new_package.last_change.parts_touched


def test_a_style_the_template_defines_is_never_replaced(new_package):
    part = new_package.style_definitions_part
    before = len(part.contents.style)
    existing = next(s for s in part.contents.style if s.style_id == "Heading1")
    existing.name.val = "My Heading One"

    ensure_style(new_package, "Heading1")

    assert len(part.contents.style) == before
    assert existing.name.val == "My Heading One"


def test_ensure_style_refuses_a_style_it_can_neither_find_nor_create(new_package):
    with pytest.raises(ContentError) as error:
        ensure_style(new_package, "NoSuchStyleAnywhere")

    assert error.value.code == "style.not_creatable"
    assert "KnownStyles" in error.value.hint


def test_the_two_code_styles_come_from_the_minimal_definitions(new_package):
    ensure_style(new_package, STYLE_IDS["code_char"])
    ensure_style(new_package, STYLE_IDS["source_code"])

    ids = style_ids_of(new_package.style_definitions_part)
    assert {"CodeChar", "SourceCode"} <= ids


# ---------------------------------------------------------------------------
# save, reload, read back
# ---------------------------------------------------------------------------


MARKDOWN = """# Report

An opening paragraph with **bold**, *italic*, ~~struck~~ and `code`.

- one
- two
  - nested

1. first
2. second

> a quotation

```
x = 1
```

| Region | Total |
| --- | --- |
| North | 120 |

See [docx4j](https://www.docx4java.org/).
"""


def test_the_whole_fragment_survives_a_save_and_a_reload(new_package):
    new_package.body.insert_markdown(MARKDOWN)
    before = new_package.to_markdown()

    after = reloaded(new_package).to_markdown()

    assert after == before
    assert after.startswith("# Report")
    assert "| Region | Total |" in after
    assert "[docx4j](https://www.docx4java.org/)" in after


def test_the_document_word_would_see_opens_as_a_valid_package(new_package):
    new_package.body.insert_markdown(MARKDOWN)
    package = reloaded(new_package)

    assert package.skipped == []
    assert package.main_document_part.skipped == []


def test_the_same_seed_and_the_same_markdown_give_the_same_bytes():
    def build() -> bytes:
        package = WordprocessingMLPackage.create_package()
        package.id_seed = 20260917
        package.body.insert_markdown(MARKDOWN)
        data = package.save()
        return data

    first, second = build(), build()
    # docProps carries the current time (CR-002 section 12.8), so compare the
    # parts the content API wrote
    assert part_bytes(first)["word/document.xml"] == part_bytes(second)["word/document.xml"]
    assert part_bytes(first)["word/numbering.xml"] == part_bytes(second)["word/numbering.xml"]
    assert part_bytes(first)["word/styles.xml"] == part_bytes(second)["word/styles.xml"]


def test_every_sample_in_the_corpus_renders_without_raising():
    for path in sorted(SAMPLES.glob("*.docx")):
        package = load(path)
        markdown = package.to_markdown()
        assert isinstance(markdown, str)
