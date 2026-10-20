#!/usr/bin/env python
"""Compare two sets of acceptance artefacts, canonically, with the dates scrubbed.

CR-007 section 4.5 point 4: the eleven artefacts of ``scripts/acceptance.py`` are
regenerated from the **installed wheel** for a release, and the record in
``tests/README.md`` says how they differ from the ones the checkout builds. This
is that comparison, the one ``tests/README.md``'s entries have been making by
hand since 2026-09-19::

    python scripts/compare_acceptance.py out/acceptance-checkout out/acceptance

For each ``.docx`` present in both directories: the zip entry lists must be
equal; a binary entry must be byte for byte equal; an XML entry is compared in
``scripts/canon.py``'s prefix-insensitive normal form after the wall-clock
dates are scrubbed --- ``dcterms:created`` and ``dcterms:modified`` in
``docProps/core.xml``, and every ``w:date`` attribute (comments, revisions,
``w:pPrChange`` and friends), which the script does not fix and Word writes as
the clock, with the ``w16du:dateUtc`` and ``w16cex:dateUtc`` beside them (CR-010
section 12.9): their values, not whether they are there. Exits 1 on any other
difference, and prints every one.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from canon import normalize  # noqa: E402

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DCTERMS = "http://purl.org/dc/terms/"

DATE_ATTRIBUTE = re.compile(rb'\s(w:date|w16du:dateUtc|w16cex:dateUtc)="[^"]*"')
DCTERMS_DATE = re.compile(
    rb"<dcterms:(created|modified)([^>]*)>[^<]*</dcterms:\1>",
)


def scrub(name: str, data: bytes) -> bytes:
    """The bytes with the wall-clock dates removed, as text, before canonicalising."""
    if name == "docProps/core.xml":
        data = DCTERMS_DATE.sub(rb"<dcterms:\1\2>DATE</dcterms:\1>", data)
    return DATE_ATTRIBUTE.sub(rb' \1="DATE"', data)


def compare_one(a: Path, b: Path) -> list[str]:
    problems: list[str] = []
    with zipfile.ZipFile(a) as za, zipfile.ZipFile(b) as zb:
        names_a = za.namelist()
        names_b = zb.namelist()
        if names_a != names_b:
            problems.append(f"entry lists differ: {sorted(set(names_a) ^ set(names_b))}")
        for name in names_a:
            if name not in names_b:
                continue
            da, db = za.read(name), zb.read(name)
            if da == db:
                continue
            if name.endswith((".xml", ".rels")):
                if normalize(scrub(name, da)) == normalize(scrub(name, db)):
                    continue
                problems.append(f"{name}: differs canonically once the dates are scrubbed")
            else:
                problems.append(f"{name}: binary entry differs ({len(da)} vs {len(db)} bytes)")
    return problems


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    if len(args) != 2:
        print(__doc__)
        return 2
    left, right = Path(args[0]), Path(args[1])
    names = sorted(p.name for p in left.glob("*.docx"))
    missing = sorted(set(names) ^ {p.name for p in right.glob("*.docx")})
    if missing:
        print(f"not in both directories: {missing}")
        return 1
    failed = 0
    compared = 0
    for name in names:
        problems = compare_one(left / name, right / name)
        with zipfile.ZipFile(left / name) as z:
            compared += len(z.namelist())
        if problems:
            failed += 1
            print(f"{name}: DIFFERS")
            for problem in problems:
                print(f"    {problem}")
        else:
            print(f"{name}: identical (canonically, dates scrubbed)")
    print(
        f"{len(names)} artefacts, {compared} entries compared, "
        f"{failed} differ beyond the wall-clock dates"
    )
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
