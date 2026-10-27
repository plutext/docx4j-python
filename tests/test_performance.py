"""Import time and parse throughput thresholds, CR-001 section 8.

These are regression guards, not benchmarks: generous thresholds that catch a
change of kind (an import that starts pulling the whole of DrawingML in eagerly,
a parse that becomes quadratic), not a change of degree. The measured figures
they are set from are in `codegen/README.md` and CR-001 section 13.

Import time is measured **cold**, in a subprocess with no `__pycache__` for the
package, because that is the number a short-lived process pays and the one
CR-001 section 8 accepted a long-lived process to avoid.

The whole module is marked ``slow`` and is **skipped under CI** (``CI=true``,
which GitHub Actions sets): the thresholds are checkout-machine numbers --- the
cold import is 0.97 s here against a 2.0 s budget --- and a CI runner is slower
by a factor nobody controls: the first run of `test.yml` (CR-007 section 10.4)
measured the cold import at 2.24 s on a GitHub runner, every other test green.
The thresholds are never loosened for a runner; a regression guard with
numbers from one machine is asserted on that machine.

    .venv-fork/bin/python -m pytest tests/test_performance.py -q
"""

from __future__ import annotations

import os
import statistics
import subprocess
import sys
import time
import zipfile
from pathlib import Path

import pytest

pytestmark = [
    pytest.mark.slow,
    pytest.mark.skipif(
        os.environ.get("CI") == "true",
        reason="checkout-machine thresholds; a CI runner measured 2.24 s against the 2.0 s budget",
    ),
]

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

#: Seconds. CR-001 section 8 measured 1.33 s cold for the baseline's 1,755
#: modules; Phase B's 164 modules do it in about a quarter of that. Anything
#: over this means the import graph has grown a new eager edge.
COLD_IMPORT_BUDGET = 2.0

#: MiB/s. REPORT.md section 4 measured 3.0 MiB/s with the lxml handler on the
#: baseline; the parent pointers cost CR-001 section 8 budgeted 20% for.
MIN_PARSE_THROUGHPUT = 0.8

MEASURE = """
import sys, time
sys.path.insert(0, {root!r})
start = time.perf_counter()
import docx4j_py.wml  # noqa: F401
print(time.perf_counter() - start)
"""


def measure_import(cache_dir: Path | None) -> float:
    """Import docx4j_py.wml in a fresh process and return the seconds.

    `PYTHONPYCACHEPREFIX` at an empty directory is what makes a run cold: the
    package's own `__pycache__` is ignored and all 164 modules are compiled
    from source, which is what a short-lived process pays on a first run.
    """
    env = dict(os.environ)
    if cache_dir is not None:
        env["PYTHONPYCACHEPREFIX"] = str(cache_dir)

    result = subprocess.run(
        [sys.executable, "-c", MEASURE.format(root=str(ROOT))],
        capture_output=True,
        text=True,
        env=env,
    )

    assert result.returncode == 0, result.stderr
    return float(result.stdout.strip().splitlines()[-1])


def test_a_cold_import_is_under_the_budget(tmp_path: Path) -> None:
    seconds = measure_import(tmp_path / "pycache")

    print(f"cold import of docx4j_py.wml: {seconds:.2f} s")
    assert seconds < COLD_IMPORT_BUDGET, f"{seconds:.2f} s"


def test_a_warm_import_is_faster_still(tmp_path: Path) -> None:
    cache = tmp_path / "pycache"
    cold = measure_import(cache)
    warm = measure_import(cache)

    print(f"import of docx4j_py.wml: cold {cold:.2f} s, warm {warm:.2f} s")
    assert warm < COLD_IMPORT_BUDGET
    assert warm <= cold


@pytest.fixture(scope="module")
def document_bytes() -> bytes:
    with zipfile.ZipFile(ROOT / "samples" / "tables.docx") as z:
        return z.read("word/document.xml")


def test_parse_throughput(document_bytes: bytes) -> None:
    from docx4j_xsdata.formats.dataclass.context import XmlContext
    from docx4j_xsdata.formats.dataclass.parsers import XmlParser
    from docx4j_xsdata.formats.dataclass.parsers.handlers import LxmlEventHandler

    from docx4j_py.runtime import parser_config

    from docx4j_py.wml import Document

    # the engine's config, without the post-bind hook: the parse alone, and link_parents apart
    config = parser_config(on_bind=None, skipped_report=False)
    parser = XmlParser(config=config, context=XmlContext(), handler=LxmlEventHandler)
    parser.from_bytes(document_bytes, Document)  # warm the class metadata

    samples = []
    for _ in range(5):
        start = time.perf_counter()
        parser.from_bytes(document_bytes, Document)
        samples.append(time.perf_counter() - start)

    mib = len(document_bytes) / (1024 * 1024)
    rate = mib / statistics.median(samples)
    print(f"parse: {rate:.1f} MiB/s over {mib * 1024:.0f} KiB")
    assert rate > MIN_PARSE_THROUGHPUT, f"{rate:.2f} MiB/s"


def test_link_parents_stays_within_its_budget(document_bytes: bytes) -> None:
    """CR-001 section 8 budgets 20% of parse for the parent pointers."""
    from docx4j_xsdata.formats.dataclass.context import XmlContext
    from docx4j_xsdata.formats.dataclass.parsers import XmlParser
    from docx4j_xsdata.formats.dataclass.parsers.handlers import LxmlEventHandler

    from docx4j_py.runtime import parser_config

    from docx4j_py.child import link_parents
    from docx4j_py.wml import Document

    # the engine's config, without the post-bind hook: the parse alone, and link_parents apart
    config = parser_config(on_bind=None, skipped_report=False)
    parser = XmlParser(config=config, context=XmlContext(), handler=LxmlEventHandler)
    doc = parser.from_bytes(document_bytes, Document)

    parse, link = [], []
    for _ in range(5):
        start = time.perf_counter()
        doc = parser.from_bytes(document_bytes, Document)
        parse.append(time.perf_counter() - start)
        start = time.perf_counter()
        link_parents(doc)
        link.append(time.perf_counter() - start)

    share = statistics.median(link) / statistics.median(parse)
    print(f"link_parents: {share:.1%} of parse")
    assert share < 0.20, f"{share:.1%}"


#: Milliseconds: ``samples/Symbols.docx``'s ``document.xml`` (790,857 bytes, 24,758 objects) parsed by the
#: engine's parser (``runtime.parser()``, as ``XmlPart._unmarshal`` parses) at ``eb8f2d5``, before CR-020 made
#: every typed list check what goes into it: the best of nine, 189.9 to 196.6 over eleven runs on the checkout
#: machine, in this test and outside it, interleaved with the code after (CR-020 section 10.9).
SYMBOLS_PARSE_BEFORE_MS = 192.0

#: CR-020 section 5: the parse within 5% of that. A parse is not checked (the check is off for its duration,
#: one context-variable set per parse, Q3), so what is left is the check's machinery in the lists the parse
#: builds: a context-variable read per non-empty list and the owner comparison in ``set_owner``.
MAX_PARSE_GROWTH = 0.05


def test_the_list_check_leaves_the_parse_within_five_percent() -> None:
    """CR-020 section 5: ``Symbols.docx``'s parse within 5% of its parse before the check (best of nine)."""
    from docx4j_py.runtime import parser
    from docx4j_py.wml import Document

    with zipfile.ZipFile(ROOT / "samples" / "Symbols.docx") as z:
        data = z.read("word/document.xml")
    parser().from_bytes(data, Document)  # warm the class metadata

    samples = []
    for _ in range(9):
        engine = parser()
        start = time.perf_counter()
        engine.from_bytes(data, Document)
        samples.append(time.perf_counter() - start)

    best = min(samples) * 1000
    limit = SYMBOLS_PARSE_BEFORE_MS * (1 + MAX_PARSE_GROWTH)
    print(f"Symbols.docx parse: {best:.1f} ms, before CR-020 {SYMBOLS_PARSE_BEFORE_MS} ms, limit {limit:.1f} ms")
    assert best < limit, f"{best:.1f} ms against {limit:.1f} ms"


#: CR-020 Phase B: the validator (``misplaced()``) over ``Symbols.docx``'s main part, against its parse. Measured at
#: 33 ms against a 195 ms parse (17%, section 11.6); the guard is a quarter, so it catches a change of kind (a
#: per-node lookup that is no longer cached, a walk that visits a node twice), not of degree.
MAX_VALIDATOR_SHARE = 0.25


def test_the_validator_costs_under_a_quarter_of_the_parse() -> None:
    """CR-020 section 11.6: ``misplaced()`` over ``Symbols.docx``'s main part, best of nine, under 25% of its parse."""
    from docx4j_py.runtime import parser
    from docx4j_py.validator import findings_in
    from docx4j_py.wml import Document

    with zipfile.ZipFile(ROOT / "samples" / "Symbols.docx") as z:
        data = z.read("word/document.xml")
    tree = parser().from_bytes(data, Document)
    findings_in(tree)  # warm the plans and the tables

    parses, walks = [], []
    for _ in range(9):
        engine = parser()
        start = time.perf_counter()
        engine.from_bytes(data, Document)
        parses.append(time.perf_counter() - start)
        start = time.perf_counter()
        assert findings_in(tree) == []
        walks.append(time.perf_counter() - start)

    parse, walk = min(parses) * 1000, min(walks) * 1000
    print(f"Symbols.docx: validator {walk:.1f} ms, parse {parse:.1f} ms ({walk / parse:.0%})")
    assert walk < parse * MAX_VALIDATOR_SHARE, f"{walk:.1f} ms against {parse:.1f} ms"
