#!/usr/bin/env bash
# Installs docx4j-python (PyPI distribution "docx4j", import name docx4j_py) into a local
# venv so the adapter is self-contained. Needs python3 3.12 or later.
set -euo pipefail
cd "$(dirname "$0")"
python3 -m venv .venv
.venv/bin/pip install --quiet docx4j==0.2.0
.venv/bin/python -I -c "import importlib.metadata as m; print('docx4j', m.version('docx4j'), '(docx4j-xsdata', m.version('docx4j-xsdata') + ')')"
