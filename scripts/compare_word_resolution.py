#!/usr/bin/env python
"""Compare Word's own Accept All and Reject All saves with this package's, structurally.

CR-010 section 12: acceptance artefact 8 comes with two copies resolved by the
engine, ``8-tracked-changes-accepted.docx`` (``revisions.accept_all``) and
``8-tracked-changes-rejected.docx`` (``revisions.reject_all``). The Word check
may save Word's own resolutions of ``8-tracked-changes.docx`` beside them, as
``out/acceptance-word/8-word-accepted.docx`` and ``8-word-rejected.docx``. This
puts each of Word's saves against ours through ``tests/revisions/description.py``
(the port of docx4j's ``DocumentDescription``, with the table, row, cell and
section properties shown), as ``tests/revisions/test_word_saves.py`` does for
the agreement fixtures::

    .venv-fork/bin/python scripts/compare_word_resolution.py
    .venv-fork/bin/python scripts/compare_word_resolution.py --ours out/acceptance --word out/acceptance-word
    .venv-fork/bin/python scripts/compare_word_resolution.py OURS.docx WORD.docx [OURS.docx WORD.docx ...]

For each pair it prints the number of description lines either side lacks
(``PortsAgreement``'s count) and, when there are any, the unified diff of the
two descriptions. Exits 0 when every pair is described alike, 1 when one
differs, 2 when a file is missing.
"""

from __future__ import annotations

import argparse
import difflib
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests" / "revisions"))

from description import describe, differing

#: Artefact 8's pairs: (ours, Word's), relative to ``--ours`` and ``--word``.
ARTEFACT_8 = (
    ("8-tracked-changes-accepted.docx", "8-word-accepted.docx"),
    ("8-tracked-changes-rejected.docx", "8-word-rejected.docx"),
)


def compare(ours: Path, word: Path, *, records: bool = True) -> tuple[int, str]:
    """How many description lines differ, and the diff (empty when none).

    Both files are loaded as saved; nothing is resolved here. ``records=False``
    leaves the table, row, cell and section properties out, as docx4j's phase 1
    comparison did.
    """
    from docx4j_py import load

    mine, theirs = describe(load(ours), records), describe(load(word), records)
    count = differing(theirs, mine)
    if not count:
        return 0, ""
    diff = difflib.unified_diff(
        theirs.splitlines(), mine.splitlines(), f"Word: {word.name}", f"ours: {ours.name}", lineterm=""
    )
    return count, "\n".join(diff)


def main(argv: list[str] | None = None) -> int:
    """Compare the pairs given, or artefact 8's two; print the counts and diffs."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("pairs", nargs="*", type=Path, help="OURS.docx WORD.docx, repeated")
    parser.add_argument("--ours", type=Path, default=ROOT / "out" / "acceptance")
    parser.add_argument("--word", type=Path, default=ROOT / "out" / "acceptance-word")
    parser.add_argument("--no-records", action="store_true", help="leave the table and section properties out")
    args = parser.parse_args(argv)

    if args.pairs:
        if len(args.pairs) % 2:
            parser.error("pairs come two at a time: OURS.docx WORD.docx")
        pairs = list(zip(args.pairs[::2], args.pairs[1::2], strict=True))
    else:
        pairs = [(args.ours / mine, args.word / theirs) for mine, theirs in ARTEFACT_8]

    missing = [str(path) for pair in pairs for path in pair if not path.is_file()]
    if missing:
        print("missing: " + ", ".join(missing))
        print("Word's saves go in out/acceptance-word/ (tests/README.md, artefact 8's row).")
        return 2

    differs = 0
    for ours, word in pairs:
        count, diff = compare(ours, word, records=not args.no_records)
        print(f"{ours.name} against Word's {word.name}: {count} lines differ")
        if count:
            differs += 1
            print(diff)
    print(f"{len(pairs)} compared, {differs} differ")
    return 1 if differs else 0


if __name__ == "__main__":
    raise SystemExit(main())
