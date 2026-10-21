"""``outline()``: what an agent reads first, and the budgets on it.

CR-003 section 3.4, Phase D.
"""

from __future__ import annotations

import json

from conftest import ROOT, sample

from docx4j_py import create_package, load
from docx4j_py.model.content import Outline, OutlineEntry
from docx4j_py.model.content.reports import DEFAULT_ENTRY_LIMIT


def test_the_outline_is_one_entry_per_block_with_an_address(report):
    outline = report.body.outline()

    assert isinstance(outline, Outline)
    assert all(isinstance(entry, OutlineEntry) for entry in outline.entries)
    assert [entry.ordinal for entry in outline.entries] == [f"body/{i}" for i in range(7)]

    first = outline.entries[0]
    assert first.kind == "paragraph"
    assert first.style_id == "Heading1"
    assert first.level == 1
    assert first.text == "Quarterly Report"
    assert first.chars == len("Quarterly Report")
    assert first.truncated is False
    assert first.address.startswith("w14:")
    assert first.para_id


def test_a_table_carries_its_shape_its_first_row_and_its_cells(report):
    table = report.body.outline().entries[4]

    assert table.kind == "table"
    assert (table.rows, table.cols) == (2, 2)
    assert table.text == "Region | Total", "the first row is the preview an agent chooses from"
    assert [child.ordinal for child in table.children] == [
        "body/4/0/0/0",
        "body/4/0/1/0",
        "body/4/1/0/0",
        "body/4/1/1/0",
    ]
    assert [child.text for child in table.children] == ["Region", "Total", "North", "120"]


def test_depth_stops_the_descent_and_says_so(report):
    outline = report.body.outline(depth=1)

    assert outline.entries[4].children == ()
    assert outline.truncated is True
    # the counts are still over everything, so an agent knows what it did not see
    assert outline.stats.paragraphs == 10
    assert outline.stats.tables == 1


def test_max_chars_cuts_the_text_and_reports_the_real_length(report):
    outline = report.body.outline(max_chars=10)
    entry = outline.entries[1]

    assert entry.text == "The quick " + "…"
    assert entry.chars == 44
    assert entry.truncated is True
    assert outline.truncated is True


def test_headings_only_is_the_table_of_contents_view(report):
    outline = report.body.outline(headings_only=True)

    assert [entry.text for entry in outline.entries] == [
        "Quarterly Report",
        "Revenue",
        "Outlook",
    ]
    assert [entry.level for entry in outline.entries] == [1, 2, 2]
    assert all(entry.children == () for entry in outline.entries)
    assert outline.stats.paragraphs == 10, "the counts are still over the whole body"


def test_the_limit_is_a_budget_that_says_when_it_bit(report):
    outline = report.body.outline(limit=3)

    assert len(outline.entries) == 3
    assert outline.truncated is True
    assert outline.stats.paragraphs == 10

    assert len(report.body.outline(limit=None).addresses()) == 11


def test_to_dict_leaves_out_what_is_empty(report):
    entry = report.body.outline().entries[1].to_dict()

    assert set(entry) == {"address", "ordinal", "para_id", "kind", "text", "chars"}
    assert "level" not in entry, "body text has no heading level"
    assert "rows" not in entry


def test_to_json_is_compact_and_reloads(report):
    outline = report.body.outline()
    loaded = json.loads(outline.to_json())

    assert '", "' not in outline.to_json(), "compact separators, because it is paid for"
    assert loaded["stats"]["paragraphs"] == 10
    assert loaded["entries"][0]["text"] == "Quarterly Report"
    assert json.loads(outline.to_json(indent=2)) == loaded


def test_to_markdown_is_a_nested_list_with_headings_as_headings(report):
    markdown = report.body.outline().to_markdown()

    assert markdown.splitlines()[0] == "# Quarterly Report"
    assert "## Revenue" in markdown
    assert "- The quick brown fox jumps over the lazy dog." in markdown
    assert "- table 2x2: Region | Total" in markdown
    assert "  - Region" in markdown, "the cells are nested under the table"


def test_the_package_outline_covers_the_headers_and_the_footers():
    package = sample("toc.docx")
    outline = package.outline()

    assert outline.entries
    assert outline.footers, "toc.docx has two footers"
    footer = outline.footers[0]
    assert footer.prefix.startswith("footer:rId")
    assert footer.part_name.startswith("/word/footer")
    assert all(entry.ordinal.startswith(footer.prefix) for entry in footer.entries)

    # and every address the outline reports resolves
    for address in outline.addresses()[:40]:
        assert package.element_at(address) is not None


def test_the_stats_count_comments_and_tracked_changes_structurally():
    package = sample("sample-docx.docx")
    stats = package.outline().stats

    assert stats.paragraphs > 0
    assert stats.tables == 1
    assert stats.words > 0
    assert stats.tracked_changes == 2, "one w:ins and one w:del in the sample"
    assert stats.comments == 0, "no comments part"
    assert stats.skipped == 0


def test_reading_an_outline_does_not_unmarshal_the_other_parts():
    package = sample("toc.docx")
    package.outline()

    settings = package.document_settings_part
    styles = package.style_definitions_part
    assert not settings.is_unmarshalled
    assert not styles.is_unmarshalled


def test_the_default_limit_is_the_documented_one():
    assert DEFAULT_ENTRY_LIMIT == 300


# ---------------------------------------------------------------------------
# on the text model (CR-011 Phase D)
# ---------------------------------------------------------------------------

PROBE = ROOT / "tests" / "fixtures" / "text-extraction" / "phase0" / "extraction-probe.docx"


def test_a_text_box_is_listed_after_its_anchor_at_its_textbox_address():
    """REQ-518, REQ-519, CR-011 Q9: after the paragraph that anchors it, flagged, read only."""
    outline = load(PROBE).outline(limit=None)
    entries = list(outline.entries)
    at = {entry.address: i for i, entry in enumerate(entries)}

    box = entries[at["body/3/textbox:0/0"]]
    assert at["body/3/textbox:0/0"] == at["body/3"] + 1
    assert (box.textbox, box.ordinal, box.text) == (True, "body/3/textbox:0/0", "CHOICE_ALPHA_123")
    assert box.to_dict()["textbox"] is True
    assert "textbox" not in entries[at["body/3"]].to_dict()
    # a box anchored in a table cell comes after the table, as the text model reads it
    assert at["body/68/0/0/0/textbox:0/0"] == at["body/68"] + 1
    assert "CHOICE_ALPHA_123" in outline.to_markdown()


def test_the_entries_text_is_the_text_models():
    """Hidden text left out (REQ-507); a deleted paragraph mark joined in the accepted view (REQ-402)."""
    package = create_package()
    package.body.insert_xml(
        "<w:p><w:r><w:t xml:space=\"preserve\">shown </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>"
    )
    entry = package.body.outline().entries[0]
    assert (entry.text, entry.chars) == ("shown ", 6)
    assert package.body.paragraphs[0].text == "shown hidden", "Office JS's text keeps it (CR-011 Q4)"

    joined = load(ROOT / "tests" / "fixtures" / "revisions" / "core-ts" / "check22" / "22-deleted.docx").body.outline()
    texts = [(entry.ordinal, entry.text) for entry in joined.entries]
    assert texts[3] == ("body/3", "Four.Five eight."), "the joined text on the first paragraph's entry"
    assert texts[4] == ("body/4", ""), "and the second's empty"
    assert joined.stats.paragraphs == 5, "the counts follow the blocks: a joined pair is one paragraph"


def test_the_counts_follow_the_text_model():
    stats = load(PROBE).outline().stats
    assert (stats.paragraphs, stats.tables) == (171, 5)
