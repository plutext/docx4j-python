"""Whole packages on eight threads, and the engine's import cost.

CR-002 section 3 ("one ``ParserConfig`` per thread, per CR-001 section 14; the
``XmlContext`` is shared") and section 7 ("import cost of the engine on top of
the model must stay under 100 ms").

    .venv-fork/bin/python -m pytest tests/openpackaging/test_threads_and_import.py -q
"""

from __future__ import annotations

import io
import re
import statistics
import subprocess
import sys
import threading
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py.openpackaging import OpcPackage, XmlPart  # noqa: E402
from docx4j_py.runtime import context  # noqa: E402

THREADS = 8
SAMPLES = sorted((ROOT / "samples").glob("*.docx"))

#: Seconds. CR-002 section 7's budget, with the model already imported.
IMPORT_BUDGET = 0.100


def load_unmarshal_save(path: Path) -> dict[str, bytes]:
    """One whole job: load, unmarshal every typed part, save, split the zip."""
    with OpcPackage.load(path) as pkg:
        pkg.unmarshal_all()
        data = pkg.save()
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def skipped_of(path: Path) -> int:
    with OpcPackage.load(path) as pkg:
        pkg.unmarshal_all()
        return len(pkg.skipped)


# ---------------------------------------------------------------------------
# threads
# ---------------------------------------------------------------------------


def test_eight_threads_give_the_same_bytes_as_a_sequential_run():
    """The whole engine, concurrently, against a shared ``XmlContext``."""
    shared = context()
    expected = {path.name: load_unmarshal_save(path) for path in SAMPLES}

    results: dict[str, dict[str, bytes]] = {}
    failures: list[BaseException] = []
    lock = threading.Lock()

    def work(path: Path) -> None:
        try:
            out = load_unmarshal_save(path)
        except BaseException as exc:  # noqa: BLE001 - the point of the test
            with lock:
                failures.append(exc)
            return
        with lock:
            results[path.name] = out

    jobs = list(SAMPLES) * 3
    with ThreadPoolExecutor(max_workers=THREADS) as pool:
        list(pool.map(work, jobs))

    assert failures == []
    assert results == expected
    assert context() is shared, "the context must still be the one process-wide one"


def test_the_skipped_report_is_per_part_under_concurrency():
    """CR-001 section 14.7: the report is the thing a shared config gets wrong.

    ``docx4j_py.runtime.parser()`` hands out a new config every call and
    ``XmlPart`` uses one per unmarshal, so every part's report is its own no
    matter how many threads are running.
    """
    expected = {path.name: skipped_of(path) for path in SAMPLES}
    assert set(expected.values()) == {0}, "the corpus must be lossless"

    wrong: list[str] = []
    lock = threading.Lock()

    def work(path: Path) -> None:
        count = skipped_of(path)
        if count != expected[path.name]:
            with lock:
                wrong.append(path.name)

    with ThreadPoolExecutor(max_workers=THREADS) as pool:
        list(pool.map(work, list(SAMPLES) * 3))
    assert wrong == []


def test_creating_packages_concurrently():
    """``create_package`` touches the shared default-styles resource."""
    from docx4j_py.openpackaging import WordprocessingMLPackage
    from docx4j_py.wml import p

    def work(index: int) -> bytes:
        pkg = WordprocessingMLPackage.create_package()
        pkg.main_document_part.contents.body.content.append(p(f"line {index}"))
        return pkg.save()

    with ThreadPoolExecutor(max_workers=THREADS) as pool:
        outputs = list(pool.map(work, range(24)))
    assert len(outputs) == 24
    assert len(set(outputs)) == 24


# ---------------------------------------------------------------------------
# import cost
# ---------------------------------------------------------------------------

MEASURE = """
import time, sys
import docx4j_py.wml            # the model, which is what the engine sits on
start = time.perf_counter()
import docx4j_py.openpackaging  # noqa: F401
print(time.perf_counter() - start)
"""


def _measure_import() -> float:
    out = subprocess.run(
        [sys.executable, "-c", MEASURE],
        capture_output=True,
        text=True,
        check=True,
        cwd=str(ROOT),
    )
    return float(out.stdout.strip().splitlines()[-1])


@pytest.mark.slow
def test_the_engine_imports_in_under_100_ms_on_top_of_the_model():
    """CR-002 section 7. Measured in a subprocess, three runs, the median.

    The engine never imports a model class at module level --- every typed part
    names its class as a *string* and resolves it on first use --- so this is
    the cost of the engine's own 20 modules, not of the model.
    """
    seconds = statistics.median(_measure_import() for _ in range(3))
    assert seconds < IMPORT_BUDGET, f"the engine took {seconds * 1000:.0f} ms to import"


def test_no_engine_module_imports_a_model_class_at_module_level():
    """The structural reason the number above is small; a cheap guard.

    ``import docx4j_py.openpackaging`` *does* end up importing the whole model,
    because ``docx4j_py/__init__.py`` carries the generator's import-order
    manifest (CR-001 section 13.5 point 3) and any import under the package runs
    it. What this pins is the engine's own half of the bargain: no module of it
    names a model class at import time, so the engine costs the model plus
    milliseconds and never the other way round.
    """
    engine = ROOT / "docx4j_py" / "openpackaging"
    offenders = []
    for path in sorted(engine.rglob("*.py")):
        for number, line in enumerate(path.read_text().splitlines(), 1):
            if line.startswith((" ", "\t", "#")):
                continue  # a deferred import inside a function is the whole point
            if re.match(
                r"(from|import)\s+docx4j_py\.(wml|dml|docprops|relationships|w15|wne|mce)\b",
                line,
            ):
                offenders.append(f"{path.relative_to(ROOT)}:{number}: {line}")
    assert offenders == []


def test_every_typed_part_names_a_model_class_that_exists():
    """The other half of the lazy-import bargain: the strings must be right."""
    from docx4j_py.openpackaging.parts import dml, docprops, wml
    from docx4j_py.openpackaging.parts.relationships_part import RelationshipsPart

    classes = [RelationshipsPart]
    for module in (wml, dml, docprops):
        for name in module.__all__:
            value = getattr(module, name)
            if isinstance(value, type) and issubclass(value, XmlPart):
                classes.append(value)

    named = 0
    for cls in classes:
        path = cls.model_class_path
        if path is None:
            continue
        named += 1
        part = cls.__new__(cls)
        assert isinstance(part.model_class, type), path
    assert named >= 20
