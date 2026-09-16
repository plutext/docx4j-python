"""The audit trail: an agent explaining itself in the document. CR-003 §3.4, §3.9.

Phase G's half of the scenario CR-003 section 3.4 calls "the single most useful
thing the API does for an AI workflow": ``pkg.author``, a comment on a range
found by ``find()``, a reply, a resolved thread, and a ``ChangeReport`` the
server hands back so the agent need not re-read the document. Phase F's tracked
changes are the other half and read the same ``pkg.author``.

Tool shaped throughout: every step here is one MCP tool call, with JSON-ready
arguments and a JSON-ready result.
"""

from __future__ import annotations

import json

from conftest import SEED, part_bytes, reloaded, sample

from docx4j_py.model.content import Author

#: What a comment thread's ``to_dict()`` must fit in, as a tool result.
THREAD_BUDGET = 2048


def audited(package):
    """``pkg.author = Author(...)``, and the comment an agent would leave."""
    package.author = Author("Claude", initials="C", email="claude@example.com")
    hit = package.find("first")[0]
    comment = hit.range(package.body).insert_comment(
        "Changed 'first' because the source document says 'red'."
    )
    return hit, comment


def test_the_audit_trail_scenario_end_to_end():
    package = sample("2010-sample1.docx")
    package.changes.clear()
    hit, comment = audited(package)

    # 1. the report is what the tool returns, and it says what was touched
    report = package.last_change
    assert report.operation == "insert_comment"
    assert report.addresses == (hit.address,)
    assert "/word/comments.xml" in report.parts_touched
    assert json.loads(report.to_json())["operation"] == "insert_comment"

    # 2. a reply and a resolution are two more calls, each with its own report
    comment.reply("Checked against the source.")
    comment.resolved = True
    assert [c.operation for c in package.changes[-2:]] == [
        "reply_to_comment",
        "resolve_comment",
    ]

    # 3. a human opens the result in Word and sees the thread; so does a reload
    back = reloaded(package).body.get_comments()
    assert len(back) == 1
    assert back[0].author_name == "Claude"
    assert back[0].initials == "C"
    assert back[0].author_email == "claude@example.com"
    assert back[0].resolved is True
    assert [r.content for r in back[0].replies] == ["Checked against the source."]
    assert back[0].get_range()[0].text == "first"


def test_a_comment_by_address_is_one_tool_call():
    package = sample("2010-sample1.docx")
    package.author = Author("Claude", initials="C")
    address = package.outline().entries[0].address
    comment = package.paragraph_at(address).insert_comment("The whole first block.")
    assert package.last_change.addresses == (address,)
    assert comment.to_dict()["address"] == address


def test_a_thread_as_a_tool_result_fits_the_budget():
    package = sample("2010-sample1.docx")
    _hit, comment = audited(package)
    comment.reply("Checked against the source.")
    comment.reply("And again.")
    out = json.dumps(comment.to_dict(), ensure_ascii=False, separators=(",", ":"))
    assert len(out.encode("utf-8")) < THREAD_BUDGET
    assert len(json.loads(out)["replies"]) == 2


def test_a_dry_run_of_insert_comment_leaves_the_package_byte_for_byte():
    """CR-003 section 14.4: the four parts a trial creates are un-created."""
    package = sample("2010-sample1.docx")
    _ = package.body  # the main part is read either way
    _ = package.style_definitions_part.contents
    before = part_bytes(package.save(), relationships=True)

    with package.dry_run() as trial:
        trial.author = Author("Claude", initials="C")
        comment = trial.body.search("first")[0].insert_comment("trial only")
        assert comment.id == 0
        assert package.get_part("/word/comments.xml") is not None

    for name in (
        "/word/comments.xml",
        "/word/commentsExtended.xml",
        "/word/commentsIds.xml",
        "/word/people.xml",
    ):
        assert package.get_part(name) is None, name
        assert package.content_type_manager.get_override_content_type(name) is None
    assert package.main_document_part.comments_part is None
    assert package.body.get_comments() == []
    assert part_bytes(package.save(), relationships=True) == before


def test_a_dry_run_over_a_document_that_already_has_comments_leaves_it_alone():
    from conftest import ROOT

    from docx4j_py import load

    package = load(ROOT / "tests" / "fixtures" / "comments-modern.docx")
    package.id_seed = SEED
    package.body.get_comments()
    # CR-003 section 12.5's first limitation: a part the real package has not
    # read is unmarshalled by the trial's first look at it, and ``styles.xml``
    # is one the comment verbs look at. Reading it here is what the real caller
    # would have done anyway.
    _ = package.style_definitions_part.contents
    before = part_bytes(package.save(), relationships=True)

    with package.dry_run() as trial:
        trial.body.paragraphs[0].insert_comment("trial only")

    assert len(package.body.get_comments()) == 1
    assert part_bytes(package.save(), relationships=True) == before


def test_the_same_seed_and_the_same_calls_give_the_same_bytes():
    def run() -> bytes:
        package = sample("2010-sample1.docx")
        _hit, comment = audited(package)
        reply = comment.reply("Checked against the source.")
        reply.resolved = True
        comment.resolved = True
        return package.save()

    first, second = run(), run()
    # the w:date of a comment is the wall clock, so compare everything else
    assert part_bytes(first).keys() == part_bytes(second).keys()
    for name, data in part_bytes(first).items():
        if name == "word/comments.xml":
            continue
        assert data == part_bytes(second)[name], name
    without_dates = [
        _strip_dates(part_bytes(x)["word/comments.xml"].decode("utf-8")) for x in (first, second)
    ]
    assert without_dates[0] == without_dates[1]


def _strip_dates(xml: str) -> str:
    import re

    return re.sub(r'w:date="[^"]*"', 'w:date=""', xml)


def test_markdown_markup_renders_the_same_comment_the_view_reads():
    from conftest import ROOT

    from docx4j_py import load

    package = load(ROOT / "tests" / "fixtures" / "comments.docx")
    comment = package.body.get_comments()[0]
    markdown = package.body.to_markdown(view="markup")
    assert f"{{>>{comment.content}<<}}" in markdown
    assert "{>>" not in package.body.to_markdown()


def test_a_comment_an_agent_inserts_shows_up_in_the_markup_view():
    package = sample("2010-sample1.docx")
    _hit, comment = audited(package)
    markdown = package.body.to_markdown(view="markup")
    assert f"{{>>{comment.content}<<}}" in markdown
