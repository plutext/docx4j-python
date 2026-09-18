#!/usr/bin/env python
"""Parse and marshal throughput, part by part. The fork's CR-001 benchmark.

    .venv-fork/bin/python scripts/bench_marshal.py
    .venv-fork/bin/python scripts/bench_marshal.py --json

Every WordprocessingML **main document part** in ``samples/`` is parsed into the
typed model and marshalled back, the median of five runs of each, and the table
is printed with the rate in MiB/s over the part's source size. ``Symbols.docx``
--- 771 KB, 24,758 elements, the document the fork's CR-001 measures --- gets a
line of its own at the end, because it is the number that CR's phases are
judged by.

The parse timed here is what :meth:`XmlPart._unmarshal` does: the fork's parser
with docx4j-python's post-bind hook, which wires the parent pointers as the
tree is built, and ``link_parents`` only when a config carries no hook (CR-002
section 12.14). The marshal is :meth:`XmlPart._marshal` itself, so the
namespace reconciliation of CR-002 section 5.6 is inside the number.
"""

from __future__ import annotations

import argparse
import json
import statistics
import sys
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

#: The document the CR's target is stated against.
HEADLINE = "Symbols.docx"

#: How many times each operation runs; the median is reported.
RUNS = 5


def documents(folder: Path) -> list[Path]:
    """Every WordprocessingML package in `folder`, by name."""
    return sorted(p for p in folder.iterdir() if p.suffix.lower() in {".docx", ".dotm"})


def measure(path: Path, runs: int) -> dict[str, Any] | None:
    """Parse and marshal `path`'s main document part, `runs` times each."""
    from docx4j_py.child import link_parents
    from docx4j_py.openpackaging.api import load
    from docx4j_py.runtime import context, parser

    package = load(path)
    part = getattr(package, "main_document_part", None)
    if part is None:
        return None
    clazz = part.model_class
    raw = part.xml  # untouched: the source bytes
    size = len(raw)

    parses = []
    obj = None
    for _ in range(runs):
        started = time.perf_counter()
        xml_parser = parser()
        obj = xml_parser.from_bytes(raw, clazz)
        if xml_parser.config.on_bind is None:
            link_parents(obj, context=context())
        parses.append(time.perf_counter() - started)

    part.set_contents(obj)
    marshals = []
    for _ in range(runs):
        started = time.perf_counter()
        out = part._marshal(obj)
        marshals.append(time.perf_counter() - started)

    parse = statistics.median(parses)
    marshal = statistics.median(marshals)
    mib = size / (1024 * 1024)
    return {
        "document": path.name,
        "part": str(part.part_name),
        "bytes": size,
        "out_bytes": len(out),
        "parse_ms": parse * 1000,
        "marshal_ms": marshal * 1000,
        "parse_mibs": mib / parse,
        "marshal_mibs": mib / marshal,
    }


def table(rows: list[dict[str, Any]]) -> str:
    """The rows as a markdown table, the headline document last and alone."""
    head = (
        f"| {'document':<38} | {'KB':>6} | {'parse ms':>8} | {'MiB/s':>6} "
        f"| {'marshal ms':>10} | {'MiB/s':>6} |"
    )
    rule = f"|{'-' * 40}|{'-' * 8}:|{'-' * 10}:|{'-' * 8}:|{'-' * 12}:|{'-' * 8}:|"
    lines = [head, rule]
    for row in rows:
        lines.append(
            f"| {row['document']:<38} | {row['bytes'] / 1024:6.0f} "
            f"| {row['parse_ms']:8.1f} | {row['parse_mibs']:6.2f} "
            f"| {row['marshal_ms']:10.1f} | {row['marshal_mibs']:6.2f} |"
        )
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    """Run the benchmark."""
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--samples", type=Path, default=ROOT / "samples")
    ap.add_argument("--runs", type=int, default=RUNS)
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    args = ap.parse_args(argv)

    from docx4j_py.runtime import warm_up

    warm_up()

    rows = []
    for path in documents(args.samples):
        row = measure(path, args.runs)
        if row is not None:
            rows.append(row)

    if args.json:
        json.dump(rows, sys.stdout, indent=2)
        sys.stdout.write("\n")
        return 0

    print(table(rows))
    headline = next((r for r in rows if r["document"] == HEADLINE), None)
    if headline is not None:
        print(
            f"\n{HEADLINE} {headline['part']}: {headline['bytes'] / 1024:.0f} KB, "
            f"parse {headline['parse_ms']:.0f} ms ({headline['parse_mibs']:.2f} MiB/s), "
            f"marshal {headline['marshal_ms']:.0f} ms ({headline['marshal_mibs']:.2f} MiB/s)"
        )
    total_parse = sum(r["parse_ms"] for r in rows)
    total_marshal = sum(r["marshal_ms"] for r in rows)
    total_kb = sum(r["bytes"] for r in rows) / 1024
    print(
        f"{len(rows)} documents, {total_kb:.0f} KB of main parts: "
        f"parse {total_parse:.0f} ms, marshal {total_marshal:.0f} ms"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
