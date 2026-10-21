#!/usr/bin/env python
"""What a kept ``mc:AlternateContent`` branch holds, with docx4j's root elements untyped and typed.

CR-012 Phase B's survey (section 5), after the 2026-10-19 prototype of its
section 3. Three commands:

    .venv-fork/bin/python scripts/mce_branches.py corpus
    .venv-fork/bin/python scripts/mce_branches.py survey [--json out.json] [ROOT ...]
    .venv-fork/bin/python scripts/mce_branches.py readers --out FILE [ROOT ...]
    .venv-fork/bin/python scripts/mce_branches.py compare-readers BEFORE AFTER

``corpus`` builds ``out/cr012-b-corpus/`` from commits, never from a working
tree: this repository's ``samples/`` and ``tests/fixtures/`` at ``HEAD``,
docx4j-core-ts's ``test/fixtures/`` at its ``HEAD`` and docx4j's
``docx4j-core-tests/src/test/resources`` at ``901a922a3``, each by ``git
archive``. ``COMMITS.txt`` there records the three commits.

``survey`` loads every WordprocessingML part of every distinct document twice:

* **untyped**, as Phases A and A2 left the parser: the branch hook
  (:func:`docx4j_py.runtime.branch_child_class`) asked about ``w:sdt``,
  ``w:tr`` and ``w:tc`` only, and no root class;
* **typed**, as this phase leaves it: the 61 roots typed in a branch by the
  container's binding, else by their root class.

For each it counts what every kept branch holds (the census: a branch's direct
children, typed by class or wildcard by name, and how deep the typing reaches),
compares the skipped reports, and saves the part, comparing each save with the
source in ``scripts/canon.py``'s form and the two saves with each other. It
exits non-zero if the typed save differs from the source where the untyped one
does not, if the two saves differ canonically, or if a skipped report differs.

``readers`` runs every reader of text or structure over every distinct corpus
document holding ``mc:AlternateContent`` in a WordprocessingML part and writes
the answers as JSON; ``compare-readers`` lists every answer that changed
between two such runs. Run it before a change and after it.
"""

from __future__ import annotations

import argparse
import collections
import contextlib
import dataclasses
import hashlib
import json
import subprocess
import sys
import time
import zipfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon

CORPUS = ROOT / "out" / "cr012-b-corpus"

#: (name, repository, commit, paths) of the corpus, ``git archive``d into CORPUS/name.
SOURCES = (
    ("docx4j-python", ROOT, "HEAD", ("samples", "tests/fixtures")),
    ("docx4j-core-ts", ROOT.parent / "docx4j-core-ts", "HEAD", ("test/fixtures",)),
    ("docx4j", ROOT.parent / "docx4j", "901a922a3", ("docx4j-core-tests/src/test/resources",)),
)

SUFFIXES = {".docx", ".docm", ".dotx", ".dotm"}

MCE_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


# ---------------------------------------------------------------------------
# the corpus
# ---------------------------------------------------------------------------


def build_corpus(target: Path = CORPUS) -> list[str]:
    """``git archive`` each source into `target`; return the ``COMMITS.txt`` lines."""
    import shutil

    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True)
    lines = []
    for name, repo, ref, paths in SOURCES:
        commit = subprocess.run(
            ["git", "-C", str(repo), "rev-parse", ref], check=True, capture_output=True, text=True
        ).stdout.strip()
        out = target / name
        out.mkdir()
        archive = subprocess.run(
            ["git", "-C", str(repo), "archive", commit, *paths], check=True, capture_output=True
        ).stdout
        subprocess.run(["tar", "-x", "-C", str(out)], input=archive, check=True)
        lines.append(f"{name} {commit} {' '.join(paths)}")
    (target / "COMMITS.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return lines


def documents(roots: list[Path]) -> list[tuple[str, Path]]:
    """Every distinct WordprocessingML package under `roots`, by content: (key, path), the first path met."""
    seen: set[str] = set()
    out: list[tuple[str, Path]] = []
    for root in roots:
        for path in sorted(p for p in root.rglob("*") if p.suffix.lower() in SUFFIXES and p.is_file()):
            digest = hashlib.sha1(path.read_bytes()).hexdigest()
            if digest in seen:
                continue
            seen.add(digest)
            try:
                key = str(path.relative_to(CORPUS))
            except ValueError:
                key = str(path)
            out.append((key, path))
    return out


def holds_alternate_content(path: Path) -> bool:
    """True when a WordprocessingML part of the package names ``mc:AlternateContent``."""
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                wml_part = name.startswith("word/") and name.endswith(".xml") and "/_rels/" not in name
                if wml_part and b"AlternateContent" in zf.read(name):
                    return True
    except (zipfile.BadZipFile, OSError):
        return False
    return False


# ---------------------------------------------------------------------------
# the two parsers
# ---------------------------------------------------------------------------


@contextlib.contextmanager
def untyped_roots():
    """Phase A2's parser: the branch hook asked about ``w:sdt``, ``w:tr`` and ``w:tc`` only, no root class.

    It replaces two of :mod:`docx4j_py.runtime`'s module functions for the
    duration: the names the hook is asked about, and step (b), the root
    class. Step (a), the container's binding, is untouched, so ``w:sdt``,
    ``w:tr`` and ``w:tc`` are bound as Phases A and A2 bind them.
    """
    from docx4j_py import runtime

    saved = runtime.wildcard_names, runtime.root_element_class
    runtime.wildcard_names = lambda: runtime.BRANCH_CHILD_NAMES
    runtime.root_element_class = lambda qname: None
    try:
        yield
    finally:
        runtime.wildcard_names, runtime.root_element_class = saved


def can_type_roots() -> bool:
    """True when this checkout's runtime has the roots (CR-012 Phase B)."""
    from docx4j_py import runtime

    return hasattr(runtime, "root_element_class")


# ---------------------------------------------------------------------------
# the census
# ---------------------------------------------------------------------------


def _is_any(value: Any) -> bool:
    from docx4j_py.child import is_any_element

    return is_any_element(value)


def _local(qname: str | None) -> str:
    return (qname or "").rpartition("}")[2]


def _label(child: Any) -> str:
    if _is_any(child):
        return "any:" + _local(child.qname)
    return "typed:" + type(child).__name__


def _branches(ac: Any) -> list[tuple[str, Any]]:
    out = [("choice", choice) for choice in ac.choice or ()]
    if ac.fallback is not None:
        out.append(("fallback", ac.fallback))
    return out


def _is_typed_ac(value: Any) -> bool:
    cls = value.__class__
    return cls.__module__ == "docx4j_py.mce" and cls.__name__ == "AlternateContent"


def _children_of(value: Any) -> list[Any]:
    """Every object below `value` one step down: its element fields, a list's items, a wildcard's children."""
    if _is_any(value):
        return [c for c in value.children or () if not isinstance(c, str)]
    if isinstance(value, (list, tuple)):
        return list(value)
    if dataclasses.is_dataclass(value) and not isinstance(value, type):
        out = []
        for f in dataclasses.fields(value):
            v = getattr(value, f.name, None)
            if v is not None and not isinstance(v, (str, int, float, bool, bytes)):
                out.append(v)
        return out
    return []


def census(root: Any) -> dict[str, collections.Counter]:
    """What every kept branch below `root` holds.

    * ``children``: a branch's direct children, ``(side, label)``;
    * ``deep``: every node below a branch, ``(side, "typed" | "any")``, the
      side the outermost branch's;
    * ``wildcard``: below a branch, each wildcard element by name;
    * ``untyped_ac``: ``mc:AlternateContent`` inside a wildcard tree.
    """
    children: collections.Counter = collections.Counter()
    deep: collections.Counter = collections.Counter()
    wildcard: collections.Counter = collections.Counter()
    untyped_ac = 0
    stack: list[tuple[Any, str | None]] = [(root, None)]
    while stack:
        node, side = stack.pop()
        if isinstance(node, (list, tuple)):
            stack.extend((item, side) for item in node)
            continue
        if side is not None:
            if _is_any(node):
                deep[(side, "any")] += 1
                wildcard[(side, _local(node.qname))] += 1
            else:
                deep[(side, "typed")] += 1
        if _is_any(node):
            if node.qname == f"{{{MCE_NS}}}AlternateContent":
                untyped_ac += 1
            stack.extend((c, side) for c in _children_of(node))
            continue
        if _is_typed_ac(node):
            for kind, branch in _branches(node):
                for child in branch.any_element or ():
                    if isinstance(child, str):
                        continue
                    children[(kind, _label(child))] += 1
                    stack.append((child, side or kind))
            continue
        stack.extend((c, side) for c in _children_of(node))
    return {
        "children": children,
        "deep": deep,
        "wildcard": wildcard,
        "untyped_ac": collections.Counter({"mc:AlternateContent": untyped_ac}),
    }


# ---------------------------------------------------------------------------
# the survey
# ---------------------------------------------------------------------------


def wml_parts(package: Any) -> list[tuple[str, Any]]:
    out = []
    for name in package.parts:
        part = package.parts[name]
        if (getattr(part, "model_class_path", None) or "").startswith("docx4j_py.wml"):
            out.append((str(name), part))
    return out


def one_pass(path: Path) -> dict[str, dict[str, Any]]:
    """Load the package, and for each WordprocessingML part: the census, the skipped report, the save."""
    from docx4j_py.openpackaging.api import load

    package = load(path)
    out = {}
    for name, part in wml_parts(package):
        source = part.xml  # never unmarshalled yet: the source bytes
        tree = part.contents
        out[name] = {
            "census": census(tree),
            "skipped": [str(item) for item in part.skipped],
            "saved": part._marshal(tree),
            "source": source,
        }
    return out


def survey(roots: list[Path], json_out: Path | None) -> int:
    typed_available = can_type_roots()
    docs = documents(roots)
    totals = {"untyped": collections.defaultdict(collections.Counter), "typed": collections.defaultdict(collections.Counter)}
    summary: collections.Counter = collections.Counter()
    problems: list[str] = []
    details: list[dict[str, Any]] = []
    started = time.perf_counter()
    for index, (key, path) in enumerate(docs, 1):
        try:
            with untyped_roots() if typed_available else contextlib.nullcontext():
                untyped = one_pass(path)
            typed = one_pass(path) if typed_available else None
        except Exception as exc:  # noqa: BLE001 - a document that cannot load is counted, not fatal
            summary["load errors"] += 1
            details.append({"document": key, "error": f"{type(exc).__name__}: {exc}"[:300]})
            continue
        summary["documents"] += 1
        for name, before in untyped.items():
            summary["parts"] += 1
            for table, counts in before["census"].items():
                totals["untyped"][table].update(counts)
            row: dict[str, Any] = {"document": key, "part": name, "skipped": len(before["skipped"])}
            identical_before = canon.compare(before["source"], before["saved"]).identical
            summary["untyped saves identical to the source"] += identical_before
            if before["skipped"]:
                summary["parts with a skipped report"] += 1
            if typed is not None:
                after = typed[name]
                for table, counts in after["census"].items():
                    totals["typed"][table].update(counts)
                rep = canon.compare(before["source"], after["saved"])
                summary["typed saves identical to the source"] += rep.identical
                if not rep.identical and identical_before:
                    problems.append(f"{key} {name}: the typed save differs from the source: {rep.differences}")
                if after["skipped"] != before["skipped"]:
                    problems.append(f"{key} {name}: the skipped report changed")
                if after["saved"] != before["saved"]:
                    summary["parts whose saves differ in bytes"] += 1
                    between = canon.compare(before["saved"], after["saved"])
                    if not between.identical:
                        problems.append(f"{key} {name}: the two saves differ: {between.differences}")
                        summary["parts whose saves differ canonically"] += 1
                    elif canon.c14n(before["saved"]) == canon.c14n(after["saved"]):
                        summary["  of which in attribute order only"] += 1
                    else:
                        summary["  of which in boolean spelling"] += 1
                    if b"AlternateContent" not in before["source"]:
                        problems.append(f"{key} {name}: the saves differ in a part with no mc:AlternateContent")
                    row["bytes_differ"] = True
                changed = after["census"]["children"] != before["census"]["children"]
                if changed:
                    summary["parts whose census changed"] += 1
                    row["census"] = {
                        "untyped": _flat(before["census"]["children"]),
                        "typed": _flat(after["census"]["children"]),
                    }
            details.append(row)
        if index % 50 == 0:
            print(f".. {index} of {len(docs)} documents, {time.perf_counter() - started:.0f} s", flush=True)

    print(f"documents {summary['documents']}, parts {summary['parts']}, load errors {summary['load errors']}")
    order = sorted(k for k in summary if k not in {"documents", "parts", "load errors"} and not k.startswith(" "))
    for key in order:
        print(f"  {key}: {summary[key]}")
        if key == "parts whose saves differ in bytes":
            for sub in sorted(k for k in summary if k.startswith(" ")):
                print(f"  {sub}: {summary[sub]}")
    print("\nwhat a kept branch holds (its direct children):")
    print(_table(totals, "children"))
    print("\nhow deep the typing reaches (every node below a branch):")
    print(_table(totals, "deep"))
    print("\nwildcard elements left below a branch, the commonest:")
    for side in ("untyped", "typed"):
        common = totals[side]["wildcard"].most_common(12)
        print(f"  {side}: " + ", ".join(f"{s}/{n} {c}" for (s, n), c in common))
    print("\nmc:AlternateContent inside a wildcard tree:", {s: totals[s]["untyped_ac"]["mc:AlternateContent"] for s in totals})
    if not typed_available:
        print("\n(this checkout's runtime has no root table: one pass, untyped)")
    if problems:
        print(f"\n{len(problems)} problem(s):")
        for line in problems[:40]:
            print("  " + line)
    if json_out is not None:
        json_out.write_text(
            json.dumps(
                {
                    "summary": dict(summary),
                    "totals": {
                        side: {table: _flat(counts) for table, counts in tables.items()}
                        for side, tables in totals.items()
                    },
                    "problems": problems,
                    "parts": details,
                },
                indent=1,
            ),
            encoding="utf-8",
        )
    return 1 if problems else 0


def _flat(counts: collections.Counter) -> dict[str, int]:
    return {" ".join(k) if isinstance(k, tuple) else str(k): v for k, v in sorted(counts.items())}


def _table(totals: dict, table: str) -> str:
    keys = sorted(set(totals["untyped"][table]) | set(totals["typed"][table]))
    lines = [f"  {'branch':9} {'child':44} {'untyped':>8} {'typed':>8}"]
    for key in keys:
        side, label = key
        lines.append(
            f"  {side:9} {label:44} {totals['untyped'][table][key]:8} {totals['typed'][table][key]:8}"
        )
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# the readers
# ---------------------------------------------------------------------------

#: The words ``find`` and ``search`` look for: common enough to hit in most documents.
WORDS = ("e", "the", "Text")


def _jsonable(value: Any) -> Any:
    return json.loads(json.dumps(value, default=repr))


def reader_answers(path: Path) -> dict[str, Any]:
    """Every reader of text or structure over one document, each answer JSON-able, or the error it raised."""
    from docx4j_py.docx import Document
    from docx4j_py.model.revisions import tracked_changes
    from docx4j_py.model.text import extract, markdown, plain_text
    from docx4j_py.openpackaging.api import load

    out: dict[str, Any] = {}

    def read(name: str, fn: Any) -> None:
        try:
            out[name] = _jsonable(fn())
        except Exception as exc:  # noqa: BLE001 - the error is the answer
            out[name] = f"ERROR {type(exc).__name__}: {exc}"[:400]

    package = load(path)
    body = package.body
    read("body.paragraphs", lambda: [[p.address, p.text] for p in body.paragraphs])
    read("body.tables", lambda: [[t.address, t.values] for t in body.tables])
    read(
        "body.inline_pictures",
        lambda: [[pic.rel_id, pic.alt_text_description, pic.width, pic.height] for pic in body.inline_pictures],
    )

    def element_at() -> list:
        addresses = [p.address for p in body.paragraphs] + [t.address for t in body.tables]
        answers = []
        for address in addresses:
            try:
                view = body.element_at(address)
                answers.append([address, type(view).__name__, getattr(view, "text", None)])
            except Exception as exc:  # noqa: BLE001
                answers.append([address, f"ERROR {type(exc).__name__}"])
        return answers

    read("body.element_at", element_at)
    read("body.outline", lambda: body.outline().to_dict())
    read("package.describe", lambda: package.describe().to_dict())
    read("body.get_text", lambda: body.get_text())
    read("body.find", lambda: {w: [h.to_dict() for h in body.find(w)] for w in WORDS})
    read("body.search", lambda: {w: [r.text for r in body.search(w)] for w in WORDS})
    read("text.extract", lambda: extract(package).to_dict())
    read("text.plain_text", lambda: plain_text(extract(package)))
    read("text.markdown", lambda: markdown(extract(package), package))
    read("package.to_markdown", lambda: package.to_markdown())
    read(
        "revisions.tracked_changes",
        lambda: [[c.type, str(c.kind), c.author, c.text] for c in tracked_changes(package)],
    )
    document = Document(str(path))
    read("docx.paragraphs", lambda: [p.text for p in document.paragraphs])
    read(
        "docx.tables",
        lambda: [[[cell.text for cell in row.cells] for row in table.rows] for table in document.tables],
    )
    read("docx.inline_shapes", lambda: [[str(s.type), s.width, s.height] for s in document.inline_shapes])
    return out


def readers(roots: list[Path], out_path: Path) -> int:
    docs = [(key, path) for key, path in documents(roots) if holds_alternate_content(path)]
    answers: dict[str, Any] = {}
    started = time.perf_counter()
    for index, (key, path) in enumerate(docs, 1):
        try:
            answers[key] = reader_answers(path)
        except Exception as exc:  # noqa: BLE001
            answers[key] = {"load": f"ERROR {type(exc).__name__}: {exc}"[:400]}
        if index % 20 == 0:
            print(f".. {index} of {len(docs)}, {time.perf_counter() - started:.0f} s", flush=True)
    out_path.write_text(json.dumps(answers, indent=1, sort_keys=True), encoding="utf-8")
    errors = sum(1 for a in answers.values() for v in a.values() if isinstance(v, str) and v.startswith("ERROR"))
    print(f"{len(docs)} documents holding mc:AlternateContent; {errors} reader errors; written to {out_path}")
    return 0


def compare_readers(before_path: Path, after_path: Path) -> int:
    before = json.loads(before_path.read_text(encoding="utf-8"))
    after = json.loads(after_path.read_text(encoding="utf-8"))
    changed: dict[str, list[str]] = collections.defaultdict(list)
    for key in sorted(set(before) | set(after)):
        a, b = before.get(key, {}), after.get(key, {})
        for reader in sorted((set(a) | set(b)) - {"trace"}):  # a traceback is the script's, not a reader's
            if a.get(reader) != b.get(reader):
                changed[reader].append(key)
    print(f"{len(before)} documents before, {len(after)} after")
    if not changed:
        print("no reader's answer changed")
        return 0
    for reader, keys in sorted(changed.items()):
        print(f"\n{reader}: {len(keys)} document(s)")
        for key in keys:
            print(f"  {key}")
            print("    before: " + _short(before.get(key, {}).get(reader)))
            print("    after:  " + _short(after.get(key, {}).get(reader)))
    return 0


def _short(value: Any, limit: int = 300) -> str:
    text = json.dumps(value, ensure_ascii=False)
    return text if len(text) <= limit else text[:limit] + " ..."


# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="command", required=True)
    sub.add_parser("corpus", help="build out/cr012-b-corpus/ from commits")
    s = sub.add_parser("survey", help="the census and the round trip, untyped and typed")
    s.add_argument("roots", nargs="*", type=Path)
    s.add_argument("--json", type=Path, default=None)
    r = sub.add_parser("readers", help="every reader's answers, as JSON")
    r.add_argument("roots", nargs="*", type=Path)
    r.add_argument("--out", type=Path, required=True)
    c = sub.add_parser("compare-readers", help="the answers that changed between two reader runs")
    c.add_argument("before", type=Path)
    c.add_argument("after", type=Path)
    args = ap.parse_args(argv)

    if args.command == "corpus":
        for line in build_corpus():
            print(line)
        return 0
    if args.command == "compare-readers":
        return compare_readers(args.before, args.after)
    roots = args.roots or [CORPUS]
    if args.command == "survey":
        return survey(roots, args.json)
    return readers(roots, args.out)


if __name__ == "__main__":
    raise SystemExit(main())
