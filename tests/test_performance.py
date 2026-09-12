"""Import time and parse throughput thresholds, CR-001 section 8.

These are regression guards, not benchmarks: generous thresholds that catch a
change of kind (an import that starts pulling the whole of DrawingML in eagerly,
a parse that becomes quadratic), not a change of degree. The measured figures
they are set from are in `codegen/README.md` and CR-001 section 13.

Import time is measured **cold**, in a subprocess with no `__pycache__` for the
package, because that is the number a short-lived process pays and the one
CR-001 section 8 accepted a long-lived process to avoid.

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
    from docx4j_xsdata.formats.dataclass.parsers.config import ParserConfig
    from docx4j_xsdata.formats.dataclass.parsers.handlers import LxmlEventHandler

    from docx4j_py.wml import Document

    config = ParserConfig(
        fail_on_unknown_properties=False,
        fail_on_unknown_attributes=False,
        fail_on_converter_warnings=False,
    )
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
    from docx4j_xsdata.formats.dataclass.parsers.config import ParserConfig
    from docx4j_xsdata.formats.dataclass.parsers.handlers import LxmlEventHandler

    from docx4j_py.child import link_parents
    from docx4j_py.wml import Document

    config = ParserConfig(
        fail_on_unknown_properties=False,
        fail_on_unknown_attributes=False,
        fail_on_converter_warnings=False,
    )
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
