#!/usr/bin/env bash
# CR-001 Phase B and C: regenerate the WordprocessingML bindings as one package
# per namespace under docx4j_py/, each with its `el` module.
#
#   codegen/generate.sh            regenerate
#   codegen/generate.sh --check    regenerate twice and prove the output is
#                                  byte identical (the import order manifest
#                                  used to shuffle, fork CHANGES.md stage 4)
#
# The generator needs the virtual environment's bin on PATH (it shells out to
# `ruff` by bare name) and it validates its own output by importing the whole
# output package. What is hand written under docx4j_py/ is listed in
# codegen/clean.py; the pre-Phase-B baseline is not there at all, it is
# baseline/docx4j_py/generated/, see baseline/docx4j_py/__init__.py.
#
# Two steps: xsdata writes one module per namespace, then codegen/generate_el.py
# turns each of those into a package of the same name and writes its el.py, the
# per-namespace object factory of CR-001 section 6.2, plus docx4j_py/el_index.py
# and the reviewed tables in codegen/el_tables/.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

VENV="$ROOT/.venv-fork"
export PATH="$VENV/bin:$PATH"

if [ "${1:-}" = "--check" ]; then
    ONE="$(mktemp -d)/one"
    TWO="$(mktemp -d)/two"
    "$0"
    cp -r docx4j_py "$ONE"
    "$0"
    cp -r docx4j_py "$TWO"
    find "$ONE" "$TWO" -name __pycache__ -type d -exec rm -rf {} + 2>/dev/null || true
    if diff -r "$ONE" "$TWO" >/dev/null; then
        echo "reproducible: two generations of docx4j_py are byte identical"
        exit 0
    fi
    echo "NOT REPRODUCIBLE"
    diff -r "$ONE" "$TWO" | head -20
    exit 1
fi

# Everything hand written under docx4j_py/ is listed here; the rest of the
# previous run is deleted first, so that a renamed namespace cannot leave a
# stale module behind.
python "$ROOT/codegen/clean.py"

# The generator takes one SOURCE, so the entry points live in one root schema:
# wml/wml.xsd and its 91-file closure, plus the four CR-002 section 9 added
# (relationships and the three docProps parts), which nothing in WML imports.
docx4j-xsdata generate -c .xsdata.phase-b.xml schemas/docx4j_python__ROOT.xsd

echo
echo "== el (CR-001 Phase C)"
python "$ROOT/codegen/generate_el.py"

echo
echo "modules: $(find docx4j_py -name '*.py' | wc -l)"
