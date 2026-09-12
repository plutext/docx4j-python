#!/usr/bin/env python
"""The thread-safety check CR-001 section 8 left open.

    .venv-fork/bin/python scripts/threads.py                 # the three runs
    .venv-fork/bin/python scripts/threads.py --iterations 20
    .venv-fork/bin/python scripts/threads.py --threads 16

Section 8 accepted the performance figures "on the condition that the library
runs in a long-lived process", and left one thing to verify: *"Thread safety of
a shared ``XmlContext`` and parser across concurrent documents is to be verified
before any concurrent server is built; until then one context per worker."*

This parses and serialises every part of every document in ``samples/``
concurrently on N threads, many times over, in three configurations, and
compares every byte against a sequential run of the same work:

``per-thread``
    one ``XmlContext`` per thread --- the conservative position section 8 fell
    back on, and the control.
``shared-context``
    **one** ``XmlContext`` for the process, a new ``XmlParser`` and a new
    ``ParserConfig`` per parse. This is what :mod:`docx4j_py.runtime` does.
``shared-parser``
    one ``XmlContext`` *and* one ``XmlParser`` for the process, which is the
    hazard: ``ParserConfig.skipped`` is a list on the config that the parser
    clears at the start of every parse and appends to during it, so two threads
    sharing a config write each other's report. Run with
    ``--skipped-report`` (the default) to see it; the run reports the damage
    rather than asserting it away.

What was read before writing this, in the fork
(``docx4j_xsdata/formats/dataclass``):

* ``context.XmlContext`` holds ``cache`` (class -> ``XmlMeta``) and
  ``xsi_cache`` (qname -> classes), plus ``sys_modules``, an int. Both caches
  only ever grow, and ``build()`` is ``if clazz not in cache: cache[clazz] =
  builder.build(...)``. Two threads racing on the same class each build the
  metadata and the last assignment wins; the metadata is derived from the class
  and is equal either way, so the race wastes work and changes nothing. There
  is no invalidation except the explicit ``clear()``, which nothing calls.
* ``parsers.bases.NodeParser`` keeps **no** per-parse state on itself: the
  ``queue`` and ``objects`` lists live on the handler, which ``parse()``
  constructs per call. The one exception is ``self.config.skipped``.
* ``serializers.XmlSerializer.render`` builds its writer per call.

Exit status is non-zero if any configuration lost output or raised.
"""

from __future__ import annotations

import argparse
import sys
import threading
import time
import traceback
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_xsdata.formats.dataclass.context import XmlContext  # noqa: E402
from docx4j_xsdata.formats.dataclass.parsers import XmlParser  # noqa: E402
from docx4j_xsdata.formats.dataclass.serializers import XmlSerializer  # noqa: E402

from docx4j_py import namespaces  # noqa: E402
from docx4j_py.child import link_parents  # noqa: E402
from docx4j_py.runtime import parser_config, serializer_config  # noqa: E402

PART_CLASSES = {
    "document.xml": "Document",
    "styles.xml": "Styles",
    "numbering.xml": "Numbering",
    "settings.xml": "CTSettings",
    "fontTable.xml": "Fonts",
    "webSettings.xml": "CTWebSettings",
    "footnotes.xml": "CTFootnotes",
    "endnotes.xml": "CTEndnotes",
    "comments.xml": "Comments",
}
HDR_FTR = "CTHdrFtr"


def load_parts() -> list[tuple[str, bytes, type]]:
    """Every WML part of every sample document, with the class it parses as."""
    import docx4j_py.wml as wml

    parts: list[tuple[str, bytes, type]] = []
    for docx in sorted((ROOT / "samples").glob("*.docx")):
        with zipfile.ZipFile(docx) as zf:
            for name in sorted(zf.namelist()):
                if not name.startswith("word/") or not name.endswith(".xml"):
                    continue
                base = name.rsplit("/", 1)[-1]
                if base.startswith(("header", "footer")):
                    class_name = HDR_FTR
                else:
                    class_name = PART_CLASSES.get(base, "")
                clazz = getattr(wml, class_name, None)
                if clazz is None:
                    continue
                parts.append((f"{docx.name}:{name}", zf.read(name), clazz))
    return parts


NS_MAP = {p: u for p, u in namespaces.PREFIXES.items() if p != "xml"}


def round_trip(
    data: bytes, clazz: type, parser: XmlParser, serializer: XmlSerializer
) -> tuple[str, int]:
    """Parse, link and serialise; also report how much the parse dropped.

    The skipped count is the interesting half: it is the one piece of per-parse
    state that lives on the *config*, so it is what a shared parser corrupts
    even when the documents themselves come out right.
    """
    obj = parser.from_bytes(data, clazz)
    link_parents(obj, context=parser.context)
    skipped = len(parser.skipped or ())
    return serializer.render(obj, NS_MAP), skipped


def make_parser(context: XmlContext, skipped_report: bool) -> XmlParser:
    return XmlParser(
        context=context, config=parser_config(skipped_report=skipped_report)
    )


def make_serializer(context: XmlContext) -> XmlSerializer:
    return XmlSerializer(context=context, config=serializer_config())


# ---------------------------------------------------------------------------


def run_sequential(parts: list, skipped_report: bool) -> dict[str, tuple[str, int]]:
    context = XmlContext()
    parser = make_parser(context, skipped_report)
    serializer = make_serializer(context)
    return {
        name: round_trip(data, clazz, parser, serializer) for name, data, clazz in parts
    }


def run_threaded(
    parts: list,
    expected: dict[str, tuple[str, int]],
    *,
    mode: str,
    threads: int,
    iterations: int,
    skipped_report: bool,
) -> dict[str, object]:
    """One configuration: run the whole corpus `iterations` times on `threads`."""
    local = threading.local()
    shared_context = XmlContext() if mode != "per-thread" else None
    shared_parser = (
        make_parser(shared_context, skipped_report) if mode == "shared-parser" else None
    )
    shared_serializer = (
        make_serializer(shared_context) if mode == "shared-parser" else None
    )

    mismatches: list[str] = []
    report_mismatches: list[str] = []
    failures: list[str] = []
    lock = threading.Lock()

    def tools() -> tuple[XmlParser, XmlSerializer]:
        if mode == "shared-parser":
            return shared_parser, shared_serializer  # type: ignore[return-value]
        if getattr(local, "parser", None) is None:
            context = shared_context if mode == "shared-context" else XmlContext()
            local.parser = make_parser(context, skipped_report)
            local.serializer = make_serializer(context)
        return local.parser, local.serializer

    def work(job: tuple[str, bytes, type]) -> None:
        name, data, clazz = job
        parser, serializer = tools()
        try:
            out, skipped = round_trip(data, clazz, parser, serializer)
        except Exception:  # noqa: BLE001 - the point of the run
            with lock:
                failures.append(f"{name}\n{traceback.format_exc()}")
            return
        want_xml, want_skipped = expected[name]
        if out != want_xml:
            with lock:
                mismatches.append(name)
        if skipped != want_skipped:
            with lock:
                report_mismatches.append(f"{name}: {skipped} against {want_skipped}")

    jobs = [job for _ in range(iterations) for job in parts]
    started = time.perf_counter()
    with ThreadPoolExecutor(max_workers=threads) as pool:
        list(pool.map(work, jobs))
    elapsed = time.perf_counter() - started

    return {
        "mode": mode,
        "jobs": len(jobs),
        "seconds": elapsed,
        "mismatches": mismatches,
        "report_mismatches": report_mismatches,
        "failures": failures,
    }


def report_race(parts: list, threads: int, rounds: int) -> tuple[int, int, str]:
    """Race the one part that loses content against one that does not.

    The corpus is 71 parts of which one drops anything, so the three runs above
    almost never put two threads in the same shared report at the same time.
    This does it on purpose: the part with the most skipped items and a part
    with none, alternating, through **one** shared ``ParserConfig``. Every read
    of ``parser.skipped`` that does not match the sequential answer is a report
    a caller would have believed.
    """
    import docx4j_py.wml as wml

    context = XmlContext()
    probe = make_parser(context, True)
    counts = []
    for name, data, clazz in parts:
        probe.from_bytes(data, clazz)
        counts.append((len(probe.skipped or ()), name, data, clazz))
    # Every part of the corpus is lossless since CR-002 Phase A's
    # CT_StylePaneFilter schema patch (toc.docx's settings.xml used to drop 15
    # attributes and was what this raced). The deliberately lossy fixture takes
    # its place, so the claim below stays tested.
    lossy = ROOT / "tests" / "fixtures" / "unknown_content.xml"
    if lossy.exists():
        counts.append((-1, lossy.name, lossy.read_bytes(), wml.Document))
        probe.from_bytes(counts[-1][2], wml.Document)
        counts[-1] = (len(probe.skipped or ()), *counts[-1][1:])
    counts.sort(key=lambda entry: -entry[0])
    worst, name, data, clazz = counts[0]
    if worst == 0:
        return 0, 0, "no part of the corpus loses content; nothing to race"
    clean_count, _clean_name, clean_data, clean_clazz = counts[-1]

    shared = make_parser(context, True)
    wrong = 0
    lock = threading.Lock()

    def work(index: int) -> None:
        nonlocal wrong
        if index % 2:
            payload, kind, want = data, clazz, worst
        else:
            payload, kind, want = clean_data, clean_clazz, clean_count
        shared.from_bytes(payload, kind)
        if len(shared.skipped or ()) != want:
            with lock:
                wrong += 1

    with ThreadPoolExecutor(max_workers=threads) as pool:
        list(pool.map(work, range(rounds)))
    return wrong, rounds, f"{name} loses {worst} item(s) sequentially"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--threads", type=int, default=8)
    ap.add_argument("--iterations", type=int, default=5)
    ap.add_argument(
        "--skipped-report",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="turn the fork's skipped-content report on; it is the shared mutable state",
    )
    ap.add_argument(
        "--modes", nargs="*", default=["per-thread", "shared-context", "shared-parser"]
    )
    args = ap.parse_args(argv)

    parts = load_parts()
    print(
        f"{len(parts)} parts from {len(list((ROOT / 'samples').glob('*.docx')))} documents"
    )
    print(
        f"{args.threads} threads, {args.iterations} iterations, skipped_report={args.skipped_report}\n"
    )

    expected = run_sequential(parts, args.skipped_report)
    print(f"sequential baseline: {len(expected)} parts\n")

    print(
        f"{'mode':<16} {'jobs':>7} {'seconds':>9} {'output diffs':>13} "
        f"{'report diffs':>13} {'failures':>9}"
    )
    bad = 0
    for mode in args.modes:
        result = run_threaded(
            parts,
            expected,
            mode=mode,
            threads=args.threads,
            iterations=args.iterations,
            skipped_report=args.skipped_report,
        )
        print(
            f"{result['mode']:<16} {result['jobs']:>7} {result['seconds']:>9.2f} "
            f"{len(result['mismatches']):>13} {len(result['report_mismatches']):>13} "
            f"{len(result['failures']):>9}"
        )
        for entry in list(result["report_mismatches"])[:3]:
            print(f"    skipped report wrong: {entry}")
        if result["mismatches"] or result["failures"]:
            bad += 1
            for name in list(result["mismatches"])[:3]:
                print(f"    output differs: {name}")
            for failure in list(result["failures"])[:2]:
                print("    " + failure.splitlines()[0])
                print("      " + failure.strip().splitlines()[-1])

    wrong, rounds, note = report_race(parts, args.threads, args.iterations * 100)
    print(f"\nskipped-report race, one shared ParserConfig ({note}):")
    print(f"  {wrong} of {rounds} reads of parser.skipped were another thread's")
    print(
        "  the parsed documents are unaffected; the report is not. One ParserConfig per thread."
    )

    print()
    if bad:
        print(f"THREAD SAFETY: {bad} configuration(s) lost output or raised")
        return 1
    print(
        "THREAD SAFETY: no configuration lost or altered a document; a shared "
        "ParserConfig loses only the skipped-content report, as above"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
