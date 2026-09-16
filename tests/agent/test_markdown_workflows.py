"""The two markdown workflows an agent has, tool-shaped. CR-003 section 3.5.

The **coarse** one is docx4j-mcp's pair: a document out as markdown, a
document in from markdown. Nothing is addressed, and whatever the markdown
cannot carry is lost --- which is the price of the cheapest possible round
trip.

The **fine** one is this CR's: ``to_markdown(addresses=True)`` to read a
document into a context window, then edits *by address*, so that only the
blocks the agent named change and every untouched part is written back byte
for byte.

Both rest on the same calls, which is the point; section 3.5's last paragraph
and the README say which to use when.
"""

from __future__ import annotations

import json
import re
import time

from conftest import SEED, big_document, part_bytes, reloaded, sample

from docx4j_py import WordprocessingMLPackage, create_package
from docx4j_py.model.markdown import ADDRESS_COMMENT

#: What an agent would write, and what a server would hand back.
BRIEF = """# Weekly brief

Three things happened, and one of them was **important**.

## What happened

- the build went green
- the corpus grew
  - two new fixtures
- a dependency was added

1. review the diff
2. run the suite
3. ship it

> "Nothing may be dropped silently."

| Area | State |
| --- | --- |
| engine | green |
| model | green |

See [the CR](https://example.invalid/cr-003) for the detail, and note `--check`.
"""


# ---------------------------------------------------------------------------
# the coarse workflow: markdown in, markdown out
# ---------------------------------------------------------------------------


def test_markdown_in_and_out_gives_back_the_same_structure():
    # ``markdown_to_docx``
    package = create_package()
    package.id_seed = SEED
    package.body.insert_markdown(BRIEF)
    data = package.save()

    # ``docx_to_markdown`` on what Word would open
    markdown = WordprocessingMLPackage.load(data).to_markdown()

    assert markdown.splitlines()[0] == "# Weekly brief"
    assert "## What happened" in markdown
    assert "- the build went green" in markdown
    assert "  - two new fixtures" in markdown
    assert "1. review the diff" in markdown
    assert "3. ship it" in markdown
    assert "> " in markdown
    assert "| Area | State |" in markdown
    assert "[the CR](https://example.invalid/cr-003)" in markdown
    assert "`--check`" in markdown


def test_the_round_trip_is_stable_a_second_time_round():
    first = create_package()
    first.id_seed = SEED
    first.body.insert_markdown(BRIEF)
    once = reloaded(first).to_markdown()

    second = create_package()
    second.id_seed = SEED
    second.body.insert_markdown(once)
    twice = reloaded(second).to_markdown()

    assert twice == once, "markdown out, markdown in, markdown out is a fixed point"


def test_the_coarse_tool_result_is_json_ready():
    package = create_package()
    package.id_seed = SEED
    package.body.insert_markdown(BRIEF)

    result = package.last_change.to_dict()
    assert json.loads(json.dumps(result))["operation"] == "insert_markdown"
    assert result["parts_touched"] == [
        "/word/document.xml",
        "/word/numbering.xml",
        "/word/styles.xml",
    ]


# ---------------------------------------------------------------------------
# the fine workflow: read with addresses, edit by address
# ---------------------------------------------------------------------------


def test_read_with_addresses_then_edit_by_one_of_them(report):
    markdown = report.to_markdown(addresses=True)

    # the agent picks a block out of the markdown with the documented regex
    blocks = ADDRESS_COMMENT.split(markdown)
    addresses = ADDRESS_COMMENT.findall(markdown)
    assert addresses, "the export emitted no addresses"
    heading = next(
        address
        for address, text in zip(addresses, blocks[1:], strict=False)
        if text.strip().startswith("## Revenue")
    )

    report.changes.clear()
    report.paragraph_at(heading).insert_paragraph("Revenue rose again.", location="After")

    change = report.last_change
    assert change.operation == "insert_paragraph"
    assert change.text_after == "Revenue rose again."
    assert len(change.addresses) == 1
    assert "Revenue rose again." in report.to_markdown()


def test_an_address_read_from_markdown_survives_a_save_and_a_reload(report):
    markdown = report.to_markdown(addresses=True)
    address = ADDRESS_COMMENT.search(markdown).group("address")

    package = reloaded(report)

    assert package.paragraph_at(address).text == "Quarterly Report"


def test_the_fine_edit_leaves_every_other_part_byte_for_byte():
    package = sample("2010-sample1.docx")
    package.body  # the only part the edit touches
    before = part_bytes(package.save())

    markdown = package.to_markdown(addresses=True)
    address = ADDRESS_COMMENT.search(markdown).group("address")
    package.paragraph_at(address).insert_text(" Edited.", location="End")
    after = part_bytes(package.save())

    changed = [name for name, data in before.items() if after.get(name) != data]
    assert changed == ["word/document.xml"]


def test_the_addresses_in_the_markdown_are_the_outlines_top_level_ones(report):
    markdown = report.to_markdown(addresses=True)

    found = [match.group("address") for match in ADDRESS_COMMENT.finditer(markdown)]
    # the outline descends into a table's cells and the markdown does not (a
    # table is one block of GFM), so ``depth=1`` is what the two agree on
    assert found == report.outline(depth=1, limit=None).addresses()


# ---------------------------------------------------------------------------
# dry run, determinism, budgets
# ---------------------------------------------------------------------------


def test_a_dry_run_of_insert_markdown_previews_without_committing():
    package = sample("2010-sample1.docx")
    package.id_seed = SEED
    package.body
    package.style_definitions_part.contents
    before = part_bytes(package.save())

    with package.dry_run() as trial:
        trial.body.insert_markdown("## Preview\n\nWhat it would say.\n")
        preview = trial.last_change.to_dict()
        assert "## Preview" in trial.to_markdown()

    after = part_bytes(package.save())
    assert [n for n, d in before.items() if after.get(n) != d] == []
    assert "## Preview" not in package.to_markdown()
    assert preview["operation"] == "insert_markdown"
    assert len(preview["addresses"]) == 2

    # and committing gives what the preview said it would
    package.body.insert_markdown("## Preview\n\nWhat it would say.\n")
    assert package.last_change.addresses == tuple(preview["addresses"])


def test_the_same_seed_and_the_same_markdown_give_the_same_bytes():
    def build() -> dict[str, bytes]:
        package = create_package()
        package.id_seed = SEED
        package.body.insert_markdown(BRIEF)
        return part_bytes(package.save())

    first, second = build(), build()
    for name, data in first.items():
        if name.startswith("docProps/"):
            continue  # the current time, CR-002 section 12.8
        assert second[name] == data, name


def test_a_200_page_document_fits_a_markdown_budget():
    package = big_document()

    whole = package.markdown_budget()
    assert whole.truncated is False
    assert whole.chars > 200_000, "a 200-page document is a lot of markdown"

    cut = package.markdown_budget(16 * 1024)
    assert cut.truncated is True
    assert cut.chars == whole.chars
    assert len(cut.text) <= 16 * 1024
    assert whole.text.startswith(cut.text), "it is cut, not summarised"
    assert cut.text.endswith("dog.") or cut.text.endswith("|"), "cut at a block boundary"

    # and the plain call returns the str, as get_text does
    assert package.to_markdown(max_chars=16 * 1024) == cut.text


def test_reading_a_200_page_document_as_markdown_is_affordable():
    package = big_document()
    package.to_markdown(max_chars=1)  # warm the numbering and style caches

    start = time.perf_counter()
    markdown = package.to_markdown()
    elapsed = time.perf_counter() - start

    assert len(markdown) > 200_000
    assert elapsed < 5.0, f"{elapsed:.2f}s to render 2,200 blocks"


def test_inserting_twenty_blocks_is_one_change_report_with_twenty_addresses():
    package = create_package()
    package.id_seed = SEED
    package.changes.clear()

    package.body.insert_markdown("\n\n".join(f"Paragraph {i}." for i in range(20)))

    assert len(package.changes) == 1
    assert len(package.last_change.addresses) == 20


def test_what_the_markdown_could_not_carry_is_in_the_report():
    package = create_package()
    package.body.insert_markdown(
        "Text.\n\n<table><tr><td>raw</td></tr></table>\n\n![remote](https://example.invalid/x.png)\n"
    )

    warnings = package.last_change.warnings
    assert any("HTML block was skipped" in w for w in warnings)
    assert any("not fetched" in w for w in warnings)
    assert json.loads(package.last_change.to_json())["warnings"] == list(warnings)


def test_the_markup_view_is_the_tracked_changes_switch_a_server_exposes():
    accepted = sample("sample-docx.docx").to_markdown(view="accepted")
    markup = sample("sample-docx.docx").to_markdown(view="markup")

    assert "{++" in markup and "{--" in markup
    assert "{++" not in accepted and "{--" not in accepted
    assert len(re.findall(r"\{\+\+.*?\+\+\}", markup)) == 1
