#!/usr/bin/env bash
# Full reproduction of the experiment from a clean checkout.
set -euo pipefail
cd "$(dirname "$0")/.."

# 1. environment -----------------------------------------------------------
uv venv --python 3.14 .venv
uv pip install --python .venv/bin/python 'xsdata[cli,lxml]==26.2'
export PATH="$PWD/.venv/bin:$PATH"      # xsdata shells out to `ruff`

.venv/bin/python -c "import sys,xsdata,lxml.etree as e; \
  print(sys.version.split()[0], xsdata.__version__, e.__version__)"

# 2. generate --------------------------------------------------------------
rm -rf docx4j_py
time .venv/bin/xsdata generate -c .xsdata.xml schemas/wml/wml.xsd

# 3. import cost -----------------------------------------------------------
find . -name __pycache__ -type d -exec rm -rf {} + 2>/dev/null || true
.venv/bin/python scripts/bench.py     # cold
.venv/bin/python scripts/bench.py     # warm

# 4. round-trip ------------------------------------------------------------
rm -rf out
# (A) as-generated defaults + a strict pass to list unknown content
.venv/bin/python scripts/roundtrip.py --strict --tolerant-factory \
    --out out/default | tee out/run-default.txt
# (B) tuned: drop schema defaults, reuse the source root's prefix bindings
.venv/bin/python scripts/roundtrip.py --ignore-defaults --nsmap-from-source \
    --tolerant-factory --out out/tuned | tee out/run-tuned.txt

# 5. targeted checks (a)-(e) ----------------------------------------------
.venv/bin/python scripts/checks.py out/tuned | tee out/checks-tuned.txt
.venv/bin/python scripts/checks.py out/default | tee out/checks-default.txt
