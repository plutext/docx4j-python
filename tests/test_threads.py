"""Concurrency, CR-001 section 8 --- the short version of ``scripts/threads.py``.

Section 8 left thread safety "to be verified before any concurrent server is
built". ``scripts/threads.py`` is the full run over ``samples/``; this is the
part that belongs in the suite: a shared :class:`XmlContext` across eight
threads gives byte-identical output, and a shared ``ParserConfig`` does not
keep its skipped-content report straight, which is why
:func:`docx4j_py.runtime.parser` hands out a new one every time.

    .venv-fork/bin/python -m pytest tests/test_threads.py -q
"""

from __future__ import annotations

import sys
import threading
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_xsdata.formats.dataclass.context import XmlContext  # noqa: E402
from docx4j_xsdata.formats.dataclass.parsers import XmlParser  # noqa: E402

from docx4j_py import namespaces as ns  # noqa: E402
from docx4j_py import wml as W  # noqa: E402
from docx4j_py.child import link_parents  # noqa: E402
from docx4j_py.runtime import parser, parser_config, serializer  # noqa: E402

THREADS = 8
NS_MAP = {p: u for p, u in ns.PREFIXES.items() if p != "xml"}


@pytest.fixture(scope="module")
def parts() -> list[tuple[str, bytes, type]]:
    """A handful of real parts, enough to make the threads collide."""
    out = []
    for name in ("tables.docx", "toc.docx", "Symbols.docx", "invoice2013.docx"):
        with zipfile.ZipFile(ROOT / "samples" / name) as zf:
            out.append((name, zf.read("word/document.xml"), W.Document))
            out.append((name, zf.read("word/styles.xml"), W.Styles))
    return out


def round_trip(data: bytes, clazz: type) -> str:
    p = parser()
    obj = p.from_bytes(data, clazz)
    link_parents(obj, context=p.context)
    return serializer().render(obj, NS_MAP)


def test_a_shared_context_gives_the_same_bytes_as_a_sequential_run(parts):
    expected = {
        name + str(len(data)): round_trip(data, clazz) for name, data, clazz in parts
    }

    failures: list[BaseException] = []
    results: dict[str, str] = {}
    lock = threading.Lock()

    def work(job):
        name, data, clazz = job
        try:
            out = round_trip(data, clazz)
        except BaseException as exc:  # noqa: BLE001 - the point of the test
            with lock:
                failures.append(exc)
            return
        with lock:
            results[name + str(len(data))] = out

    jobs = parts * 4
    with ThreadPoolExecutor(max_workers=THREADS) as pool:
        list(pool.map(work, jobs))

    assert failures == []
    assert results == expected


def test_one_parser_config_per_thread_is_required(parts):
    """The one piece of shared mutable state, and what it costs.

    ``ParserConfig.skipped`` is a list the parser clears at the start of every
    parse and appends to during it. Two threads through one config therefore
    read each other's report --- the documents are fine, the report is not.
    The fixture has to be a part that actually loses content; ``toc.docx``'s
    ``settings.xml`` was that part until CR-002's ``CT_StylePaneFilter`` schema
    patch made it lossless, so ``tests/fixtures/unknown_content.xml`` --- which
    is deliberately lossy and always will be --- takes its place.
    """
    lossy = (ROOT / "tests" / "fixtures" / "unknown_content.xml").read_bytes()
    clean = parts[0][1]

    probe = parser()
    probe.from_bytes(lossy, W.Document)
    want_lossy = len(probe.skipped)
    assert want_lossy > 0, "the fixture must actually lose something"
    probe.from_bytes(clean, W.Document)
    assert len(probe.skipped) == 0

    shared = XmlParser(context=XmlContext(), config=parser_config())
    wrong = 0
    lock = threading.Lock()

    def work(index: int) -> None:
        nonlocal wrong
        data, want = (lossy, want_lossy) if index % 2 else (clean, 0)
        clazz = W.Document
        shared.from_bytes(data, clazz)
        if len(shared.skipped) != want:
            with lock:
                wrong += 1

    with ThreadPoolExecutor(max_workers=THREADS) as pool:
        list(pool.map(work, range(400)))

    # Not asserted to be non-zero: it is a race and a run may win it. What is
    # asserted is that a parser of one's own never gets it wrong.
    own = []
    with ThreadPoolExecutor(max_workers=THREADS) as pool:
        own = list(
            pool.map(
                lambda i: _skipped_with_own_parser(
                    lossy if i % 2 else clean, i % 2, want_lossy
                ),
                range(200),
            )
        )
    assert all(own), (
        f"a per-thread config must always be right (shared got {wrong} wrong)"
    )


def _skipped_with_own_parser(data: bytes, lossy: int, want_lossy: int) -> bool:
    p = parser()
    p.from_bytes(data, W.Document)
    return len(p.skipped) == (want_lossy if lossy else 0)
