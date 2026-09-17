#!/usr/bin/env bash
# Reproduce REPORT.md's experiment: upstream xsdata 26.2, the `clusters` layout.
#
# The output is the pre-Phase-B **baseline**, `baseline/docx4j_py/generated/`,
# which is git-ignored and derived. Nothing here touches `docx4j_py/`, the
# hand-written engine and the fork-generated model; `codegen/generate.sh` owns
# that. xsdata writes its package relative to the working directory, which is
# why step 2 runs inside `baseline/`.
set -euo pipefail
cd "$(dirname "$0")/.."

# 1. environment -----------------------------------------------------------
if [ ! -x .venv/bin/python ]; then
  uv venv --python 3.14 .venv
  uv pip install --python .venv/bin/python 'xsdata[cli,lxml]==26.2'
fi
export PATH="$PWD/.venv/bin:$PATH"      # xsdata shells out to `ruff`

.venv/bin/python -c "import sys,xsdata,lxml.etree as e; \
  print(sys.version.split()[0], xsdata.__version__, e.__version__)"

# 2. generate the baseline -------------------------------------------------
rm -rf baseline/docx4j_py
mkdir -p baseline
( cd baseline && time ../.venv/bin/xsdata generate -c ../.xsdata.xml ../schemas/wml/wml.xsd )
test -f baseline/docx4j_py/generated/__init__.py

# 3. import cost -----------------------------------------------------------
find baseline -name __pycache__ -type d -exec rm -rf {} + 2>/dev/null || true
.venv/bin/python scripts/bench.py     # cold
.venv/bin/python scripts/bench.py     # warm

# 4. round-trip ------------------------------------------------------------
# roundtrip.py exits non-zero on any real difference (a rule added after the
# experiment); config A is *expected* to differ (REPORT.md section 7), so the
# exit codes here are reported, not fatal.
rm -rf out/default out/tuned
# (A) as-generated defaults + a strict pass to list unknown content
( .venv/bin/python scripts/roundtrip.py --strict --tolerant-factory \
    --out out/default || echo "config A: roundtrip.py exited $?" ) 2>&1 | tee out/run-default.txt
# (B) tuned: drop schema defaults, reuse the source root's prefix bindings
( .venv/bin/python scripts/roundtrip.py --ignore-defaults --nsmap-from-source \
    --tolerant-factory --out out/tuned || echo "config B: roundtrip.py exited $?" ) 2>&1 | tee out/run-tuned.txt

# 5. targeted checks (a)-(e) ----------------------------------------------
( .venv/bin/python scripts/checks.py out/tuned || echo "checks (tuned) exited $?" ) 2>&1 | tee out/checks-tuned.txt
( .venv/bin/python scripts/checks.py out/default || echo "checks (default) exited $?" ) 2>&1 | tee out/checks-default.txt
