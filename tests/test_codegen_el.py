"""The ``el`` output is reproducible and matches its committed tables.

CR-001 section 13.4 pinned the generated model as byte identical over two runs;
Phase C adds the ``el`` modules, the ``el_index`` and the tables to that
promise. The cheap half runs always; the whole regeneration is ``slow``.

    .venv-fork/bin/python -m pytest tests/test_codegen_el.py -q
    .venv-fork/bin/python -m pytest tests/test_codegen_el.py -q -m "not slow"
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "codegen"))

from generate_el import identifier, snake  # noqa: E402

from docx4j_py.el_index import EL_MODULES  # noqa: E402

TABLES = ROOT / "codegen" / "el_tables"


def test_a_table_is_committed_for_every_el_module():
    names = {path.stem for path in TABLES.glob("*.json")}
    modules = {
        m.removeprefix("docx4j_py.").removesuffix(".el").replace(".", "_")
        for m in EL_MODULES.values()
    }
    assert names == modules


def test_every_table_reads_and_says_which_class_each_name_takes():
    from importlib import import_module

    for uri, module_name in EL_MODULES.items():
        name = (
            module_name.removeprefix("docx4j_py.").removesuffix(".el").replace(".", "_")
        )
        table = json.loads((TABLES / f"{name}.json").read_text(encoding="utf-8"))
        assert table["namespace"] == uri
        assert table["module"] == module_name.removesuffix(".el")
        module = import_module(module_name)
        assert set(table["elements"]) == set(module.ELEMENTS)
        for local, class_name in table["elements"].items():
            assert module.ELEMENTS[local].__name__ == class_name
        assert table["counts"]["element_names"] == len(module.ELEMENTS)
        assert "overrides" in table


def test_the_naming_rules_the_tables_record():
    assert snake("CTSdtRow") == "sdt_row"
    assert snake("SdtBlock") == "sdt_block"
    assert snake("PHyperlink") == "p_hyperlink"
    assert snake("CTCustomXmlRun") == "custom_xml_run"
    assert snake("CTFFName") == "ff_name"
    assert snake("RT") == "rt"

    assert identifier("p") == "p"
    assert identifier("pPr") == "pPr"  # the element name, verbatim
    assert identifier("del") == "del_"  # a keyword
    assert identifier("type") == "type_"  # a builtin
    assert identifier("id") == "id_"
    assert identifier("name") == "name"  # neither


@pytest.mark.slow
def test_regenerating_el_changes_nothing():
    """`codegen/generate_el.py` is a function of the metadata and the overrides."""
    before = {path: path.read_bytes() for path in sorted(TABLES.glob("*.json"))}
    before |= {
        path: path.read_bytes() for path in sorted(ROOT.glob("docx4j_py/**/el.py"))
    }
    before[ROOT / "docx4j_py" / "el_index.py"] = (
        ROOT / "docx4j_py" / "el_index.py"
    ).read_bytes()

    result = subprocess.run(
        [sys.executable, str(ROOT / "codegen" / "generate_el.py"), "--no-convert"],
        capture_output=True,
        text=True,
        cwd=ROOT,
    )
    assert result.returncode == 0, result.stderr

    changed = [
        str(path.relative_to(ROOT))
        for path, data in before.items()
        if path.read_bytes() != data
    ]
    assert changed == []


# ---------------------------------------------------------------------------
# what ``codegen/clean.py`` must not delete (CR-003 Phase K)
# ---------------------------------------------------------------------------


def test_every_hand_written_path_is_covered_by_cleans_keep():
    """`codegen/clean.py` deletes what ``KEEP`` does not name, so ``KEEP`` must
    name everything hand written.

    The generator owns exactly three shapes of file: each namespace package's
    ``__init__.py`` and ``el.py``, the root ``docx4j_py/__init__.py`` and
    ``docx4j_py/el_index.py``. Every *other* tracked path under ``docx4j_py/``
    is hand written and must be inside ``KEEP``, or the next
    ``codegen/generate.sh`` silently deletes it --- which is what would have
    happened to ``docx4j_py/model/`` before this phase added it.

    Read from git rather than from the working tree, and asserted without
    running the generator: ``generate.sh`` is not to be run from a test.
    """
    from clean import KEEP

    tracked = subprocess.run(
        ["git", "ls-files", "docx4j_py"],
        capture_output=True,
        text=True,
        cwd=ROOT,
        check=True,
    ).stdout.split()
    assert tracked, "git listed no files under docx4j_py/"

    kept = {Path(entry) for entry in KEEP}
    generated = {"el.py", "el_index.py"}

    uncovered: list[str] = []
    for entry in tracked:
        relative = Path(entry).relative_to("docx4j_py")
        if relative in kept or any(parent in kept for parent in relative.parents):
            continue
        if relative.name in generated or relative.name == "__init__.py":
            continue  # the generator writes this one
        uncovered.append(entry)

    assert uncovered == [], (
        "these hand-written paths are not in codegen/clean.py's KEEP and the "
        f"next regeneration would delete them: {uncovered}"
    )


def test_everything_keep_names_is_really_there():
    """A stale ``KEEP`` entry hides a path that no longer exists."""
    from clean import KEEP

    missing = [entry for entry in sorted(KEEP) if not (ROOT / "docx4j_py" / entry).exists()]
    assert missing == []
