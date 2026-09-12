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
