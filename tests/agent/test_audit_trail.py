"""The audit trail, end to end: author, tracking, an edit, a comment. CR-003 §3.4.

Phase F's half of "the audit trail is Word's own", over Phase G's. One scripted
session of tool-shaped calls --- set the identity, turn tracking on, edit by
address, count a replacement in a ``dry_run`` and then make it, explain it in a
comment, hand the caller the changes as JSON, save --- and then the reload a
human's Word would see, where accepting everything gives exactly what the same
edits made with tracking off give.
"""

from __future__ import annotations

import datetime
import json

from conftest import SEED, part_bytes, reloaded, sample

from docx4j_py.model.content import Author

#: The date the session fixes, so that the saved bytes are reproducible.
WHEN = datetime.datetime(2026, 9, 17, 12, 0, tzinfo=datetime.UTC)

#: The document the session edits: a Word 2010 file with no comment parts,
#: no ``w14:paraId`` and no ``w:trackRevisions``.
SOURCE = "2010-sample1.docx"


def session(*, tracking: bool = True):
    """A package with the agent's identity, its seed and its date fixed."""
    package = sample(SOURCE)
    package.author = Author("Claude", initials="C", email="claude@example.invalid")
    package.tracked_change_date = WHEN
    if tracking:
        package.change_tracking_mode = "TrackAll"
    return package


def edit(package) -> int:
    """The edits themselves, so that the tracked and untracked runs are the same calls."""
    hit = package.find("first")[0]
    paragraph = package.paragraph_at(hit.address)
    paragraph.insert_text("Reviewed. ", location="Start")
    count = package.body.replace_text("document", "report")
    package.body.insert_paragraph("Checked against the source.", style="Normal")
    return count


def test_the_audit_trail_is_the_documents_own():
    package = session()
    assert package.change_tracking_mode == "TrackAll"
    assert package.author.name == "Claude"

    # what the replacement would do, before doing it
    with package.dry_run() as trial:
        previewed = trial.body.replace_text("document", "report")
    assert previewed >= 1

    count = edit(package)
    assert count == previewed

    # and the agent says why, in the document
    comment = package.find("report")[0].range(package.body).insert_comment(
        "Changed 'document' to 'report' because the brief says report."
    )
    assert comment.author_name == "Claude"

    changes = package.get_tracked_changes()
    assert changes, "every edit since the mode went on is a revision"
    assert {c.author for c in changes} == {"Claude"}
    assert {c.date for c in changes} == {WHEN}

    first = changes[0].to_dict()
    assert set(first) == {"type", "author", "date", "text", "id", "kind", "address"}
    assert first["author"] == "Claude"
    assert json.loads(json.dumps(changes[0].to_dict()))["date"] == "2026-09-17T12:00:00+00:00"

    # a tool would hand back the report of the last call, and the change list
    assert package.last_change.operation == "insert_comment"
    assert "/word/comments.xml" in package.last_change.parts_touched

    back = reloaded(package)
    assert len(back.get_tracked_changes()) == len(changes)
    assert [c.content for c in back.body.get_comments()] == [comment.content]


def test_accepting_everything_gives_what_an_untracked_run_gives():
    tracked = session()
    edit(tracked)
    plain = session(tracking=False)
    edit(plain)

    assert tracked.body.text == plain.body.text, "the accepted view already matches"

    back = reloaded(tracked)
    accepted = back.body.accept_all()
    assert accepted == len(tracked.get_tracked_changes())
    assert back.body.text == plain.body.text
    assert "w:ins" not in back.body.get_xml() and "w:del" not in back.body.get_xml()


def test_rejecting_everything_gives_the_document_back():
    package = session()
    before = package.body.text
    edit(package)
    assert package.body.text != before

    package.body.reject_all()
    assert package.body.text == before


def test_the_comment_survives_accepting_and_rejecting():
    package = session()
    package.find("first")[0].range(package.body).insert_text("New. ", location="Before")
    comment = package.find("first")[0].range(package.body).insert_comment("Why this changed.")

    accepted = reloaded(package)
    accepted.body.accept_all()
    assert [c.content for c in accepted.body.get_comments()] == [comment.content]

    rejected = reloaded(package)
    rejected.body.reject_all()
    assert [c.content for c in rejected.body.get_comments()] == [comment.content]


def test_the_same_seed_the_same_date_and_the_same_calls_give_the_same_bytes():
    """CR-003 section 12.6, with ``tracked_change_date`` standing in for the clock."""
    first = session()
    edit(first)
    first.find("report")[0].range(first.body).insert_comment("Why this changed.")

    second = session()
    edit(second)
    second.find("report")[0].range(second.body).insert_comment("Why this changed.")

    one, two = part_bytes(first.save()), part_bytes(second.save())
    assert set(one) == set(two)
    for name in one:
        if name == "word/comments.xml":
            # the comment's own w:date is the wall clock (CR-003 section 15.2)
            assert _masked(one[name]) == _masked(two[name])
            continue
        assert one[name] == two[name], name


def test_only_the_parts_the_session_touched_are_rewritten():
    package = session()
    before = part_bytes(sample(SOURCE).save())
    edit(package)
    after = part_bytes(package.save())

    changed = sorted(name for name, data in before.items() if after.get(name) != data)
    assert changed == ["word/document.xml", "word/settings.xml", "word/styles.xml"]
    assert package.id_seed == SEED


def _masked(data: bytes) -> bytes:
    """The bytes with every ``w:date`` value blanked."""
    import re

    return re.sub(rb'w:date="[^"]*"', b'w:date=""', data)
