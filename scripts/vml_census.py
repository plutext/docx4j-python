#!/usr/bin/env python
"""VML in every WordprocessingML part, before and after a change: the census, the skipped reports, the saves.

CR-012 Phase C's survey (section 4.3 item 5), a sibling of
``scripts/mce_branches.py``, whose corpus it reads (``out/cr012-b-corpus/``,
built from commits by ``mce_branches.py corpus``). Phase B could compare an
untyped and a typed parse in one run, by switching the branch hook off; VML is
typed by the model itself, so this survey compares two runs, one made before
the change and one after it:

    .venv-fork/bin/python scripts/vml_census.py run --out out/cr012-c-survey/before [ROOT ...]
    .venv-fork/bin/python scripts/vml_census.py run --out out/cr012-c-survey/after [ROOT ...]
    .venv-fork/bin/python scripts/vml_census.py compare out/cr012-c-survey/before out/cr012-c-survey/after

``run`` loads every distinct document (by content) and, for each
WordprocessingML part (document, headers, footers, notes, comments, settings,
the glossary and the rest), records:

* **the census**: every ``v:``, ``o:``, ``w10:``, ``x:`` (``xvml``) and
  ``pvml:`` element, typed or wildcard, in an ``mc:Choice``, an
  ``mc:Fallback`` or outside any branch, by part type; and every
  ``w:txbxContent``, ``w:pict`` and ``w:object`` the same way;
* the skipped report;
* the save (the part unmarshalled and marshalled), compared with the source in
  ``scripts/canon.py``'s form. The saves go into ``saves.zip``, so that
  ``compare`` can set the two runs' saves side by side.

``compare`` exits non-zero on a newly skipped item, or a part whose save is
canonically identical to its source before and is not after (or whose
difference categories grew), and lists every part whose two saves differ in
bytes, by kind, with every attribute whose value changed.
"""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import sys
import time
import zipfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon  # noqa: E402
from mce_branches import CORPUS, documents  # noqa: E402

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: The VML vocabularies, by the prefix ``docx4j_py.namespaces`` gives them.
VML_PREFIXES = {
    "urn:schemas-microsoft-com:vml": "v",
    "urn:schemas-microsoft-com:office:office": "o",
    "urn:schemas-microsoft-com:office:word": "w10",
    "urn:schemas-microsoft-com:office:excel": "xvml",
    "urn:schemas-microsoft-com:office:powerpoint": "pvml",
}

#: The WordprocessingML elements VML lives in, or holds, counted alongside it.
HOLDERS = {f"{{{W}}}pict": "w:pict", f"{{{W}}}object": "w:object", f"{{{W}}}txbxContent": "w:txbxContent"}


def _ns(qname: str | None) -> str:
    return qname[1:].partition("}")[0] if qname and qname.startswith("{") else ""


def _local(qname: str | None) -> str:
    return (qname or "").rpartition("}")[2]


def census(root: Any) -> dict[str, collections.Counter]:
    """Every VML element below `root`, and every holder, by (prefix:local, typed|any, where).

    ``where`` is ``choice`` or ``fallback`` for an element in a kept branch
    (the outermost branch's side), else ``outside``. Walked with
    :func:`docx4j_py.child.iter_children` (``mce="all"``, wildcards
    included), so a typed object is named as the serialiser writes it.
    """
    from docx4j_py.child import is_any_element, iter_children

    elements: collections.Counter = collections.Counter()
    stack: list[tuple[Any, str]] = [(root, "outside")]
    while stack:
        node, where = stack.pop()
        for qname, child in iter_children(node, mce="all", wildcards=True):
            if isinstance(child, str) or child is None:
                continue
            here = where
            if qname in (f"{{{MC}}}Choice", f"{{{MC}}}Fallback") and where == "outside":
                here = "choice" if qname.endswith("Choice") else "fallback"
            kind = "any" if is_any_element(child) else "typed"
            prefix = VML_PREFIXES.get(_ns(qname))
            if prefix is not None:
                elements[(f"{prefix}:{_local(qname)}", kind, here)] += 1
            elif qname in HOLDERS:
                elements[(HOLDERS[qname], kind, here)] += 1
            stack.append((child, here))
    return {"elements": elements}


def wml_parts(package: Any) -> list[tuple[str, Any]]:
    out = []
    for name in package.parts:
        part = package.parts[name]
        if (getattr(part, "model_class_path", None) or "").startswith("docx4j_py.wml"):
            out.append((str(name), part))
    return out


def _flat(counts: collections.Counter) -> dict[str, int]:
    return {" ".join(k) if isinstance(k, tuple) else str(k): v for k, v in sorted(counts.items())}


def run(roots: list[Path], out: Path) -> int:
    from docx4j_py.openpackaging.api import load

    out.mkdir(parents=True, exist_ok=True)
    docs = documents(roots)
    records: list[dict[str, Any]] = []
    totals: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)
    summary: collections.Counter = collections.Counter()
    started = time.perf_counter()
    with zipfile.ZipFile(out / "saves.zip", "w", zipfile.ZIP_DEFLATED) as saves:
        for index, (key, path) in enumerate(docs, 1):
            try:
                package = load(path)
                parts = wml_parts(package)
            except Exception as exc:  # noqa: BLE001 - a document that cannot load is counted, not fatal
                summary["load errors"] += 1
                records.append({"document": key, "error": f"{type(exc).__name__}: {exc}"[:300]})
                continue
            summary["documents"] += 1
            for name, part in parts:
                summary["parts"] += 1
                source = part.xml  # never unmarshalled yet: the source bytes
                row: dict[str, Any] = {"document": key, "part": name, "type": type(part).__name__}
                try:
                    tree = part.contents
                    saved = part._marshal(tree)
                except Exception as exc:  # noqa: BLE001
                    summary["part errors"] += 1
                    row["error"] = f"{type(exc).__name__}: {exc}"[:300]
                    records.append(row)
                    continue
                counted = census(tree)["elements"]
                for (label, kind, where), count in counted.items():
                    totals[row["type"]][(label, kind, where)] += count
                rep = canon.compare(source, saved)
                row.update(
                    {
                        "source_sha1": hashlib.sha1(source).hexdigest(),
                        "saved_sha1": hashlib.sha1(saved).hexdigest(),
                        "identical": rep.identical,
                        "counts": dict(sorted(rep.counts.items())),
                        "skipped": [str(item) for item in part.skipped],
                        "census": _flat(counted),
                    }
                )
                summary["saves identical to the source"] += rep.identical
                summary["parts with VML"] += any(k[0].split(":")[0] in VML_PREFIXES.values() for k in counted)
                summary["parts with a skipped report"] += bool(row["skipped"])
                saves.writestr(f"{key}::{name}", saved)
                records.append(row)
            if index % 50 == 0:
                print(f".. {index} of {len(docs)} documents, {time.perf_counter() - started:.0f} s", flush=True)
    (out / "census.json").write_text(
        json.dumps(
            {
                "summary": dict(summary),
                "totals": {ptype: _flat(counts) for ptype, counts in sorted(totals.items())},
                "parts": records,
            },
            indent=1,
        ),
        encoding="utf-8",
    )
    print(f"documents {summary['documents']}, parts {summary['parts']}, load errors {summary['load errors']}, "
          f"part errors {summary['part errors']}, {time.perf_counter() - started:.0f} s")
    for key in ("saves identical to the source", "parts with VML", "parts with a skipped report"):
        print(f"  {key}: {summary[key]}")
    print_census(totals)
    return 0


def print_census(totals: dict[str, Any]) -> None:
    """The census by vocabulary, kind and place, then by part type."""
    def split(key: Any) -> tuple[str, str, str]:
        return tuple(key) if isinstance(key, tuple) else tuple(key.split(" "))  # type: ignore[return-value]

    by_prefix: collections.Counter = collections.Counter()
    by_part: collections.Counter = collections.Counter()
    holders: collections.Counter = collections.Counter()
    for ptype, counts in totals.items():
        for key, count in counts.items():
            label, kind, where = split(key)
            prefix = label.split(":")[0]
            if label in HOLDERS.values():
                holders[(label, kind, where)] += count
                continue
            by_prefix[(prefix, kind, where)] += count
            by_part[(ptype, kind)] += count
    print("\nVML elements, by vocabulary, kind and place:")
    for (prefix, kind, where), count in sorted(by_prefix.items()):
        print(f"  {prefix:5} {kind:6} {where:9} {count:7}")
    print("\nVML elements, by part type and kind:")
    for (ptype, kind), count in sorted(by_part.items()):
        print(f"  {ptype:28} {kind:6} {count:7}")
    print("\nthe holders:")
    for (label, kind, where), count in sorted(holders.items()):
        print(f"  {label:15} {kind:6} {where:9} {count:7}")


# ---------------------------------------------------------------------------
# compare
# ---------------------------------------------------------------------------


def _attr_changes(before: bytes, after: bytes) -> collections.Counter:
    """Each attribute whose value differs between two saves, by element, attribute and the two values' shapes."""
    rep = canon.DiffReport()
    canon.diff_trees(canon.normalize_tree(before), canon.normalize_tree(after), rep)
    out: collections.Counter = collections.Counter()
    for diff in rep.diffs:
        element = diff.path.rpartition("/")[2]
        out[(diff.category, element, diff.detail if diff.category != "text" else "text")] += 1
    return out


def byte_kind(before: bytes, after: bytes) -> str:
    """How two saves that differ in bytes differ."""
    rep = canon.compare(before, after)
    if not rep.identical:
        return "canonical: " + ", ".join(f"{k} {v}" for k, v in sorted(rep.differences.items()))
    if rep.counts:
        return "boolean spelling (" + ", ".join(f"{k} {v}" for k, v in sorted(rep.counts.items())) + ")"
    if canon.c14n(before) == canon.c14n(after):
        return "attribute order or namespace declarations only (equal in C14N)"
    return "prefixes only (equal in canon's normal form)"


def compare(before_dir: Path, after_dir: Path) -> int:
    before = json.loads((before_dir / "census.json").read_text(encoding="utf-8"))
    after = json.loads((after_dir / "census.json").read_text(encoding="utf-8"))
    rows_b = {(r["document"], r.get("part")): r for r in before["parts"]}
    rows_a = {(r["document"], r.get("part")): r for r in after["parts"]}
    problems: list[str] = []
    kinds: collections.Counter = collections.Counter()
    by_type: collections.Counter = collections.Counter()
    changes: collections.Counter = collections.Counter()
    examples: dict[str, str] = {}
    summary: collections.Counter = collections.Counter()
    for key in sorted(set(rows_b) | set(rows_a), key=str):
        b, a = rows_b.get(key), rows_a.get(key)
        if b is None or a is None:
            problems.append(f"{key}: in one run only")
            continue
        if "error" in b or "error" in a:
            if b.get("error") != a.get("error"):
                problems.append(f"{key}: error before {b.get('error')!r}, after {a.get('error')!r}")
            continue
        summary["parts"] += 1
        new_skips = [s for s in a["skipped"] if s not in b["skipped"]]
        if new_skips:
            problems.append(f"{key}: {len(new_skips)} newly skipped, e.g. {new_skips[0][:200]}")
        if len(a["skipped"]) < len(b["skipped"]):
            summary["parts whose skipped report shrank"] += 1
        if b["identical"] and not a["identical"]:
            problems.append(f"{key}: a new canonical difference: {a['counts']}")
        elif not a["identical"]:
            real_b = {k: v for k, v in b["counts"].items() if k not in canon.DiffReport.BENIGN}
            real_a = {k: v for k, v in a["counts"].items() if k not in canon.DiffReport.BENIGN}
            if any(real_a.get(k, 0) > real_b.get(k, 0) for k in real_a):
                problems.append(f"{key}: its canonical differences grew: {real_b} -> {real_a}")
        if not b["identical"] and a["identical"]:
            summary["parts canonically identical now, not before"] += 1
        if b["saved_sha1"] != a["saved_sha1"]:
            summary["parts whose saves differ in bytes"] += 1
            with zipfile.ZipFile(before_dir / "saves.zip") as zb, zipfile.ZipFile(after_dir / "saves.zip") as za:
                name = f"{key[0]}::{key[1]}"
                saved_b, saved_a = zb.read(name), za.read(name)
            kind = byte_kind(saved_b, saved_a)
            kinds[kind] += 1
            by_type[(a["type"], kind)] += 1
            examples.setdefault(kind, f"{key[0]} {key[1]}")
            if not kind.startswith("attribute order") and not kind.startswith("prefixes"):
                changes.update(_attr_changes(saved_b, saved_a))
    print(f"{summary['parts']} parts in both runs")
    for k in sorted(summary):
        if k != "parts":
            print(f"  {k}: {summary[k]}")
    print("\nthe byte differences, by kind:")
    for kind, count in kinds.most_common():
        print(f"  {count:5}  {kind}\n         e.g. {examples[kind]}")
    print("\nby part type:")
    for (ptype, kind), count in sorted(by_type.items()):
        print(f"  {ptype:24} {count:5}  {kind[:90]}")
    if changes:
        print("\nwhat changed, between the two saves (category, element, detail), the commonest 60:")
        for (category, element, detail), count in changes.most_common(60):
            print(f"  {count:6}  {category:18} {element:28} {detail[:110]}")
    print("\nthe census, before:")
    print_census(before["totals"])
    print("\nthe census, after:")
    print_census(after["totals"])
    if problems:
        print(f"\n{len(problems)} problem(s):")
        for line in problems[:60]:
            print("  " + line)
        return 1
    print("\nno newly skipped item, no new canonical difference")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="command", required=True)
    r = sub.add_parser("run", help="the census, the skipped reports and the saves of one checkout")
    r.add_argument("roots", nargs="*", type=Path)
    r.add_argument("--out", type=Path, required=True)
    c = sub.add_parser("compare", help="two runs side by side")
    c.add_argument("before", type=Path)
    c.add_argument("after", type=Path)
    args = ap.parse_args(argv)
    if args.command == "compare":
        return compare(args.before, args.after)
    return run(args.roots or [CORPUS], args.out)


if __name__ == "__main__":
    raise SystemExit(main())
