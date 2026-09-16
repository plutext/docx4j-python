"""``DocumentSession``: many documents, many threads, one document at a time.

CR-003 section 3.4 and decided question 5; CR-001 section 14.7 for the thread
rule this rests on.
"""

from __future__ import annotations

import threading
import time
from concurrent.futures import ThreadPoolExecutor

import pytest
from conftest import SAMPLES, WITHOUT_PARA_IDS

from docx4j_py import DocumentSession, WordprocessingMLPackage, create_package
from docx4j_py.model.content import ContentError


def test_open_get_save_close():
    session = DocumentSession()
    handle = session.open(SAMPLES / WITHOUT_PARA_IDS)

    assert isinstance(handle, str)
    assert 6 <= len(handle) <= 16, "a short opaque string, not a path"
    assert handle in session
    assert session.handles() == [handle]

    package = session.get(handle)
    assert isinstance(package, WordprocessingMLPackage)
    package.body.insert_paragraph("added through a session")

    data = session.save(handle)
    assert isinstance(data, bytes), "no target is the bytes; the source needs overwrite=True"
    assert "added through a session" in WordprocessingMLPackage.load(data).body.text

    session.close(handle)
    assert handle not in session
    session.close(handle)  # idempotent


def test_save_writes_back_over_what_was_opened(tmp_path):
    path = tmp_path / "round.docx"
    path.write_bytes((SAMPLES / WITHOUT_PARA_IDS).read_bytes())

    with DocumentSession() as session:
        handle = session.open(path)
        session.get(handle).body.insert_paragraph("edited in place")
        assert session.save(handle, overwrite=True) is None

        with pytest.raises(ContentError) as raised:
            session.add(create_package())
            session.save(session.handles()[-1], overwrite=True)
        assert raised.value.code == "session.no_source"

    assert "edited in place" in WordprocessingMLPackage.load(path).body.text


def test_add_puts_a_created_package_in_the_session():
    with DocumentSession() as session:
        handle = session.add(create_package())
        session.get(handle).body.insert_paragraph("Hello")
        assert isinstance(session.save(handle), bytes)


def test_an_unknown_handle_says_what_to_do():
    session = DocumentSession()
    with pytest.raises(ContentError) as raised:
        session.get("nope")
    assert raised.value.code == "session.unknown_handle"
    assert "handles()" in raised.value.hint
    assert "swept" in raised.value.hint


def test_max_open_refuses_and_says_so():
    with DocumentSession(max_open=1) as session:
        session.open(SAMPLES / WITHOUT_PARA_IDS)
        with pytest.raises(ContentError) as raised:
            session.open(SAMPLES / WITHOUT_PARA_IDS)
        assert raised.value.code == "session.full"
        assert "close(handle)" in raised.value.hint


def test_sweep_closes_what_has_been_idle_and_nothing_else():
    with DocumentSession(idle_timeout=0.05) as session:
        stale = session.open(SAMPLES / WITHOUT_PARA_IDS)
        time.sleep(0.08)
        fresh = session.open(SAMPLES / WITHOUT_PARA_IDS)

        closed = session.sweep()
        assert closed == [stale]
        assert session.handles() == [fresh]
        assert session.sweep() == []

    # a timeout of zero turns it off
    with DocumentSession(idle_timeout=0) as session:
        session.open(SAMPLES / WITHOUT_PARA_IDS)
        assert session.sweep() == []
        assert len(session) == 1


def test_exit_closes_everything():
    session = DocumentSession()
    with session:
        session.open(SAMPLES / WITHOUT_PARA_IDS)
        session.open(SAMPLES / "tables.docx")
        assert len(session) == 2
    assert len(session) == 0


def test_documents_are_json_ready():
    with DocumentSession() as session:
        handle = session.open(SAMPLES / WITHOUT_PARA_IDS)
        session.get(handle).body.insert_paragraph("x")
        row = session.documents()[0].to_dict()
        assert row["handle"] == handle
        assert row["source"].endswith(WITHOUT_PARA_IDS)
        assert row["changes"] == 1
        assert row["idle_for"] >= 0


def test_eight_threads_hammering_two_handles():
    """CR-003 Phase D: concurrent tool calls on one document serialise.

    Eight threads, two documents, two hundred edits each. The per-handle lock
    is what makes the counts add up; ``XmlPart`` taking a fresh ``ParserConfig``
    per unmarshal (CR-001 section 14.7) is what makes the loads safe.
    """
    edits = 25
    with DocumentSession() as session:
        handles = [
            session.open(SAMPLES / WITHOUT_PARA_IDS),
            session.open(SAMPLES / "tables.docx"),
        ]
        failures: list[BaseException] = []
        seen: dict[str, set[int]] = {handle: set() for handle in handles}
        guard = threading.Lock()

        def worker(index: int) -> None:
            handle = handles[index % 2]
            try:
                for step in range(edits):
                    with session.use(handle) as package:
                        body = package.body
                        before = len(body)
                        paragraph = body.insert_paragraph(f"t{index} s{step}")
                        # inside the lock the document is this thread's alone
                        assert len(body) == before + 1
                        assert package.last_change.addresses == (paragraph.address,)
                        with guard:
                            seen[handle].add(len(body))
            except BaseException as error:  # noqa: BLE001 - reported below
                failures.append(error)

        with ThreadPoolExecutor(max_workers=8) as pool:
            list(pool.map(worker, range(8)))

        assert not failures, failures
        for handle in handles:
            package = session.get(handle)
            added = [p.text for p in package.body.paragraphs if p.text.startswith("t")]
            assert len(added) == 4 * edits, "four threads per document, every edit landed"
            assert len(set(added)) == len(added), "and none was lost to a race"
            # the lengths seen inside the lock are consecutive: no two threads
            # ever saw the same one
            assert len(seen[handle]) == 4 * edits


def test_one_parser_config_per_parse_is_what_xml_part_already_does():
    """CR-001 section 14.7's rule, asserted where a session would rely on it."""
    from docx4j_py.runtime import context, parser

    first, second = parser(), parser()
    assert first is not second
    assert first.config is not second.config, "the skipped-content report is per parse"
    assert context() is context(), "one XmlContext per process, which is safe to share"
