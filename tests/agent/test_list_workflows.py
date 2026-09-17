"""Lists as an MCP server would use them. CR-003 section 3.10, Phase H.

Tool-shaped, like the rest of ``tests/agent/``: every step is one call with
JSON-ready arguments and a JSON-ready result --- ``outline``, ``find``, an edit
by address, the ``ChangeReport`` checked, a ``dry_run`` before committing, and
the tracked edit a human reviews in Word.
"""

from __future__ import annotations

import datetime

from conftest import SEED, part_bytes, reloaded, sample

from docx4j_py.model.content import Author

WHEN = datetime.datetime(2026, 9, 17, 12, 0, tzinfo=datetime.UTC)


def test_an_agent_reads_the_outline_then_makes_a_list_by_address():
    package = sample("2010-sample1.docx")

    # 1. outline: what is there, and the addresses to work with
    outline = package.outline().to_dict()
    addresses = [entry["ordinal"] for entry in outline["entries"]]
    assert addresses[:2] == ["body/0", "body/1"]

    # 2. start a list on the first block, by address
    first = package.paragraph_at("body/0")
    the_list = first.start_new_list()
    report = package.last_change.to_dict()
    assert report["operation"] == "start_new_list"
    assert f"list:{the_list.id}" in report["addresses"]
    assert "/word/numbering.xml" in report["parts_touched"]

    # 3. attach the next two blocks to it, one of them a level deeper
    package.paragraph_at("body/1").attach_to_list(the_list.id)
    package.paragraph_at("body/2").attach_to_list(the_list.id, 1)
    assert package.last_change.to_dict()["text_after"] == f"list {the_list.id} level 1"

    # 4. read it back the way a tool result would report it
    items = [p.to_dict() for p in package.body.paragraphs if p.is_list_item]
    assert [item["list_item"]["list_string"] for item in items] == ["1.", "2.", "a."]
    assert [item["list_item"]["level"] for item in items] == [0, 0, 1]

    # 5. and after a save and a reload it is the same document
    back = reloaded(package)
    assert [p.list_item.list_string for p in back.body.paragraphs if p.is_list_item] == [
        "1.",
        "2.",
        "a.",
    ]


def test_find_then_attach_the_paragraph_the_hit_is_in():
    package = sample("2010-sample1.docx")
    package.id_seed = SEED
    body = package.body
    for text in ("Check the invoice", "Check the licence", "Ship it"):
        body.insert_paragraph(text)
    the_list = body.paragraph_at(contains="Check the invoice").start_new_list()

    # the server's ``find`` tool: hits with addresses, no second call per hit
    hits = [hit.to_dict() for hit in body.find("Check the licence", limit=1)]
    assert [hit["snippet"] for hit in hits] == ["Check the licence"]

    package.paragraph_at(hits[0]["address"]).attach_to_list(the_list.id)
    assert [p.list_item.list_string for p in body.paragraphs if p.is_list_item] == ["1.", "2."]


def test_a_dry_run_previews_the_list_and_the_commit_gives_the_same_report():
    package = sample("2010-sample1.docx")
    package.id_seed = SEED
    _ = package.body, package.style_definitions_part.contents
    before = part_bytes(package.save(), relationships=True)

    with package.dry_run() as trial:
        trial.paragraph_at("body/0").start_new_list()
        preview = trial.last_change.to_dict()
        assert trial.body.paragraphs[0].to_dict()["list_item"]["list_string"] == "1."

    after = part_bytes(package.save(), relationships=True)
    assert [n for n, d in before.items() if after.get(n) != d] == []
    assert not package.body.paragraphs[0].is_list_item

    package.paragraph_at("body/0").start_new_list()
    committed = package.last_change.to_dict()
    assert committed["operation"] == preview["operation"]
    assert committed["addresses"] == preview["addresses"]
    assert sorted(committed["parts_touched"]) == sorted(preview["parts_touched"])
    assert package.body.paragraphs[0].list_item.list_string == "1."


def test_a_tracked_attach_is_a_formatting_revision_a_reject_undoes():
    package = sample("2010-sample1.docx")
    package.id_seed = SEED
    package.author = Author("Claude", initials="C")
    package.tracked_change_date = WHEN
    first = package.body.paragraphs[0]
    the_list = first.start_new_list()

    package.change_tracking_mode = "TrackAll"
    target = package.body.paragraphs[1]
    before = target.get_xml()
    target.attach_to_list(the_list.id)

    changes = [change.to_dict() for change in package.get_tracked_changes()]
    assert [change["type"] for change in changes] == ["Formatted"]
    assert changes[0]["author"] == "Claude"

    back = reloaded(package)
    assert back.body.reject_all() == 1
    rejected = back.body.paragraphs[1]
    assert not rejected.is_list_item
    assert "w:numPr" not in rejected.get_xml()
    assert rejected.get_xml() == before, "byte for byte the paragraph that was there"
    assert back.body.paragraphs[0].is_list_item, "and the list itself is untouched"


def test_the_report_of_every_list_verb_names_the_parts_it_will_re_marshal():
    package = sample("2010-sample1.docx")
    package.id_seed = SEED
    body = package.body
    the_list = body.paragraphs[0].start_new_list()

    operations = []
    body.paragraphs[1].attach_to_list(the_list.id)
    operations.append(package.last_change)
    the_list.set_level_numbering(0, "LowerLetter", "%1)")
    operations.append(package.last_change)
    the_list.insert_paragraph("A new item", "End")
    operations.append(package.last_change)
    body.paragraphs[1].detach_from_list()
    operations.append(package.last_change)

    assert [report.operation for report in operations] == [
        "attach_to_list",
        "list.set_level_numbering",
        "list.insert_paragraph",
        "detach_from_list",
    ]
    assert all("/word/document.xml" in report.parts_touched for report in operations)
    numbering = [
        report.operation
        for report in operations
        if "/word/numbering.xml" in report.parts_touched
    ]
    assert numbering == ["list.set_level_numbering"], (
        "only the call that wrote a definition re-marshals the numbering part"
    )
