#!/usr/bin/env python
"""Parent pointers over the sample corpus: correctness, and what they cost.

CR-001 section 5 and 8: after ``XmlParser`` returns, ``link_parents`` wires the
whole tree, and the budget for it is 20% of the parse time.

    .venv-fork/bin/python scripts/parents.py
    .venv-fork/bin/python scripts/parents.py --only tables.docx --repeat 5

For every part it

* parses it (median of N),
* runs ``link_parents`` (median of N) and reports the ratio,
* checks that *every* node's parent is the object that actually holds it, both
  after the walk and after ``deep_copy`` of a subtree,
* checks that the parsed tree still serialises to the same bytes with the
  pointers in place, i.e. the slot is invisible to the serialiser.
"""

from __future__ import annotations

import argparse
import statistics
import sys
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from roundtrip import NS_MAP, load_models, part_class

from docx4j_py.child import (
    Child,
    ChildList,
    _child_field_names,
    deep_copy,
    link_parents,
)


def check(root: object, context: object) -> tuple[int, list[str]]:
    """Verify every node's parent. Returns (node count, list of complaints).

    The expected parent is worked out from what each object's fields actually
    hold, never from the ``parent`` slot, which is the thing under test.
    """
    problems: list[str] = []
    stack: list[tuple[object, object | None]] = [(root, None)]
    seen = {id(root)}
    count = 0
    while stack:
        obj, want = stack.pop()
        count += 1
        got = obj.get_parent() if isinstance(obj, Child) else None
        if got is not want:
            problems.append(
                f"{type(obj).__name__}: parent is "
                f"{type(got).__name__ if got is not None else None}, "
                f"expected {type(want).__name__ if want is not None else None}"
            )
        for name in _child_field_names(obj.__class__, context):
            value = getattr(obj, name, None)
            if value is None:
                continue
            if isinstance(value, ChildList) and value.owner is not obj:
                problems.append(
                    f"{type(obj).__name__}.{name}: ChildList owner is wrong"
                )
            items = value if isinstance(value, (list, tuple)) else (value,)
            for item in items:
                if isinstance(item, Child) and id(item) not in seen:
                    seen.add(id(item))
                    stack.append((item, obj))
    return count, problems


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", default=None)
    ap.add_argument("--samples", default=str(ROOT / "samples"))
    ap.add_argument("--repeat", type=int, default=3)
    ap.add_argument("--models-module", default="docx4j_py.wml")
    ap.add_argument("--package", default="docx4j_py")
    args = ap.parse_args()

    models = load_models(args.package, args.models_module)

    from docx4j_xsdata.formats.dataclass.context import XmlContext
    from docx4j_xsdata.formats.dataclass.parsers import XmlParser
    from docx4j_xsdata.formats.dataclass.parsers.config import ParserConfig
    from docx4j_xsdata.formats.dataclass.parsers.handlers import LxmlEventHandler
    from docx4j_xsdata.formats.dataclass.serializers import XmlSerializer
    from docx4j_xsdata.formats.dataclass.serializers.config import SerializerConfig
    from docx4j_xsdata.formats.dataclass.serializers.writers import LxmlEventWriter

    ctx = XmlContext()
    parser = XmlParser(
        config=ParserConfig(
            fail_on_unknown_properties=False,
            fail_on_unknown_attributes=False,
            fail_on_converter_warnings=False,
        ),
        context=ctx,
        handler=LxmlEventHandler,
    )
    serializer = XmlSerializer(
        config=SerializerConfig(xml_declaration=True, indent=None),
        context=ctx,
        writer=LxmlEventWriter,
    )

    total_parse = total_link = total_nodes = 0.0
    failures = 0
    rows = []

    docs = sorted(Path(args.samples).glob("*.docx"))
    if args.only:
        docs = [d for d in docs if args.only in d.name]

    for docx in docs:
        with zipfile.ZipFile(docx) as z:
            for name in sorted(
                n for n in z.namelist() if n.startswith("word/") and n.endswith(".xml")
            ):
                klass = part_class(name, models)
                if klass is None:
                    continue
                data = z.read(name)

                parse_ms = []
                link_ms = []
                for _ in range(args.repeat):
                    t = time.perf_counter()
                    obj = parser.from_bytes(data, klass)
                    parse_ms.append((time.perf_counter() - t) * 1000)
                    t = time.perf_counter()
                    link_parents(obj, context=ctx)
                    link_ms.append((time.perf_counter() - t) * 1000)

                p = statistics.median(parse_ms)
                q = statistics.median(link_ms)
                count, problems = check(obj, ctx)

                # a copied subtree must be self consistent and detached
                copied = deep_copy(obj)
                _, copy_problems = check(copied, ctx)
                if copied.get_parent() is not None:
                    copy_problems.append("deep_copy root has a parent")

                # the slot must not reach the serializer
                before = serializer.render(
                    parser.from_bytes(data, klass), ns_map=dict(NS_MAP)
                )
                after = serializer.render(obj, ns_map=dict(NS_MAP))
                if before != after:
                    problems.append("serialisation changed once parents were linked")

                bad = problems + [f"copy: {p_}" for p_ in copy_problems]
                failures += len(bad)
                total_parse += p
                total_link += q
                total_nodes += count
                rows.append((docx.name, name, count, p, q, len(bad)))
                flag = "OK " if not bad else "BAD"
                print(
                    f"{flag} {docx.name:38s} {name:26s} nodes={count:6d} "
                    f"parse={p:7.2f}ms link={q:6.2f}ms "
                    f"{q / p * 100 if p else 0:5.1f}%"
                )
                for b in bad[:5]:
                    print(f"      {b}")

    print(f"\nparts {len(rows)}  nodes {total_nodes:.0f}")
    print(
        f"parse total {total_parse:.0f} ms, link_parents total {total_link:.0f} ms, "
        f"{total_link / total_parse * 100:.1f}% of parse "
        f"(CR-001 section 8 budget: 20%)"
    )
    print(f"parent problems: {failures}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
