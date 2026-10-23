#!/usr/bin/env bash
# CR-007 section 6 and RELEASING.md "Gates": the release gates that the suite on the
# checkout cannot provide, scripted so that a release is this script and the Word
# check rather than a memory of what was done last time.
#
#   scripts/release_check.sh                 # everything below, about four minutes
#   RELEASE_CHECK_DIR=/some/dir scripts/release_check.sh   # keep the scratch venvs there
#
# What it does, in order, stopping at the first failure:
#
#   1. builds the sdist and the wheel with `python -m build` in a scratch virtual
#      environment (uv, CPython 3.12) into dist/, and runs `twine check --strict`;
#   2. lists the wheel (`unzip -l`) and the sdist, and asserts what CR-007 section
#      4.7 says ships: docx4j_py/resources/** (the XML and JSON), py.typed, LICENSE
#      and NOTICE present; no __pycache__, nothing from codegen/, scripts/ or
#      baseline/ in the wheel; tests/, samples/, schemas/, codegen/ and scripts/ in
#      the sdist and no docx4j_py/**/__pycache__;
#   3. makes a clean CPython 3.12 virtual environment with nothing in it, installs
#      the wheel FROM dist/ with its dependencies resolved from PyPI (so
#      docx4j-xsdata comes from PyPI, never from a checkout), plus pytest, and
#      prints where docx4j_py and docx4j_xsdata resolve to;
#   4. runs a smoke script from a directory outside the repository: warm_up(),
#      load a corpus document, edit through the content API (a heading, which adds a
#      style from resources/KnownStyles.xml; markdown, which creates numbering.xml
#      from resources/openpackaging/), the fonts (resources/fonts/), save, reload, the
#      python-docx facade, create_package() --- every packaged resource is read;
#   5. runs the fast suite AGAINST THE INSTALLED WHEEL: the sdist is unpacked in the
#      scratch directory and its docx4j_py/ removed, so the tests' own
#      `sys.path.insert(0, ROOT)` finds no package there and `import docx4j_py`
#      resolves to site-packages (asserted). tests/test_codegen_el.py is skipped:
#      it tests the checkout (git ls-files, codegen/clean.py against the tree);
#   6. regenerates out/acceptance/ from the installed wheel (scripts/acceptance.py
#      copied beside a copy of samples/, so its own `sys.path.insert(0, ROOT)` finds
#      no checkout) and compares the twelve artefacts against ones built from the
#      checkout with .venv-fork (scripts/compare_acceptance.py: canonically, dates
#      scrubbed). The Word check of out/acceptance/ is the user's, recorded in
#      tests/README.md.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

UV="${UV:-uv}"
PYTHON_VERSION="${RELEASE_CHECK_PYTHON:-3.12}"
SCRATCH="${RELEASE_CHECK_DIR:-$(mktemp -d -t docx4j-release-XXXXXX)}"
mkdir -p "$SCRATCH"
CHECKOUT_PYTHON="${DOCX4J_PY_VENV:-$ROOT/.venv-fork}/bin/python"

step() { echo; echo "== $*"; }
fail() { echo "FAILED: $*" >&2; exit 1; }

command -v "$UV" >/dev/null || fail "uv is not installed (https://docs.astral.sh/uv/)"

# ---------------------------------------------------------------- 1. build
step "1. build the sdist and the wheel (scratch venv, CPython $PYTHON_VERSION)"
BUILD_VENV="$SCRATCH/build"
"$UV" venv -q --clear -p "$PYTHON_VERSION" "$BUILD_VENV"
"$UV" pip install -q --python "$BUILD_VENV/bin/python" build twine
VERSION="$("$BUILD_VENV/bin/python" -c 'import tomllib; print(tomllib.load(open("pyproject.toml", "rb"))["project"]["version"])')"
NAME="$("$BUILD_VENV/bin/python" -c 'import tomllib; print(tomllib.load(open("pyproject.toml", "rb"))["project"]["name"])')"
echo "$NAME $VERSION, $("$BUILD_VENV/bin/python" --version), $("$BUILD_VENV/bin/python" -m build --version | head -1), twine $("$BUILD_VENV/bin/twine" --version | tail -1)"
rm -rf "$ROOT/dist"
"$BUILD_VENV/bin/python" -m build --outdir "$ROOT/dist" . >"$SCRATCH/build.log" 2>&1 || { cat "$SCRATCH/build.log"; fail "python -m build"; }
WHEEL="$ROOT/dist/${NAME}-${VERSION}-py3-none-any.whl"
SDIST="$ROOT/dist/${NAME}-${VERSION}.tar.gz"
[ -f "$WHEEL" ] || fail "no wheel at $WHEEL"
[ -f "$SDIST" ] || fail "no sdist at $SDIST"
ls -l "$ROOT/dist"
"$BUILD_VENV/bin/twine" check --strict "$ROOT"/dist/*

# ---------------------------------------------------------------- 2. contents
step "2. what the wheel carries"
unzip -l "$WHEEL" | grep -v "docx4j_py/[a-z0-9_]*/__init__.py$\|docx4j_py/[a-z0-9_]*/el.py$" | sed 's/^/    /'
echo "    (the namespace packages' __init__.py and el.py elided: $(unzip -l "$WHEEL" | grep -c "docx4j_py/[a-z0-9_]*/__init__.py$\|docx4j_py/[a-z0-9_]*/el.py$") of them)"
WHEEL_LIST="$(unzip -Z1 "$WHEEL")"
WHEEL_FILES="$(echo "$WHEEL_LIST" | wc -l)"
WHEEL_MODULES="$(echo "$WHEEL_LIST" | grep -c '\.py$')"
echo "wheel: $WHEEL_FILES files, $WHEEL_MODULES modules, $(stat -c %s "$WHEEL") bytes"

must_have() { echo "$WHEEL_LIST" | grep -qx "$1" || fail "wheel lacks $1"; }
must_have "docx4j_py/py.typed"
must_have "docx4j_py/resources/KnownStyles.xml"
must_have "docx4j_py/resources/warmup_document.xml"
must_have "docx4j_py/resources/warmup_styles.xml"
must_have "docx4j_py/resources/openpackaging/styles.xml"
must_have "docx4j_py/resources/openpackaging/numbering.xml"
must_have "docx4j_py/resources/openpackaging/fontTable.xml"
must_have "docx4j_py/resources/fonts/substitutions.json"
must_have "docx4j_py/resources/fonts/families.json"
must_have "docx4j_py/resources/fonts/registry.json"
must_have "docx4j_py/resources/fonts/theme-2007.xml"
must_have "docx4j_py/resources/fonts/theme-2013.xml"
must_have "docx4j_py/resources/fonts/theme-2023.xml"
must_have "${NAME}-${VERSION}.dist-info/licenses/LICENSE"
must_have "${NAME}-${VERSION}.dist-info/licenses/NOTICE"
echo "$WHEEL_LIST" | grep -q "__pycache__" && fail "wheel carries __pycache__"
echo "$WHEEL_LIST" | grep -q "\.pyc$" && fail "wheel carries .pyc"
echo "$WHEEL_LIST" | grep -q "^codegen/\|^scripts/\|^baseline/\|^tests/\|^samples/\|^schemas/" && fail "wheel carries something from outside docx4j_py/"
echo "$WHEEL_LIST" | grep -qv "^docx4j_py/\|^${NAME}-${VERSION}.dist-info/" && fail "wheel carries an unexpected top-level entry: $(echo "$WHEEL_LIST" | grep -v "^docx4j_py/\|^${NAME}-${VERSION}.dist-info/")"
# every git-tracked file under docx4j_py/ is in the wheel
while read -r tracked; do
    echo "$WHEEL_LIST" | grep -qx "$tracked" || fail "wheel lacks tracked $tracked"
done < <(git ls-files docx4j_py)
echo "every one of $(git ls-files docx4j_py | wc -l) git-tracked files under docx4j_py/ is in the wheel; LICENSE, NOTICE, py.typed and the 18 resources present; no __pycache__"

step "2. what the sdist carries"
SDIST_LIST="$(tar -tzf "$SDIST" | sed "s|^${NAME}-${VERSION}/||")"
echo "$SDIST_LIST" | grep -v "^docx4j_py/\|^tests/\|^schemas/\|^codegen/\|^samples/\|^docs/\|^scripts/" | sed 's/^/    /'
for top in docx4j_py tests samples schemas codegen scripts docs; do
    printf "    %-12s %5d entries\n" "$top/" "$(echo "$SDIST_LIST" | grep -c "^$top/")"
done
echo "sdist: $(echo "$SDIST_LIST" | wc -l) entries, $(stat -c %s "$SDIST") bytes"
for needed in docx4j_py/py.typed tests/test_readme.py tests/README.md samples/sample-docx.docx schemas/docx4j_python__ROOT.xsd codegen/generate.sh codegen/clean.py scripts/acceptance.py scripts/canon.py docs/pypi-readme.md .xsdata.phase-b.xml pyproject.toml README.md CHANGELOG.md RELEASING.md LICENSE NOTICE PKG-INFO; do
    echo "$SDIST_LIST" | grep -qx "$needed" || fail "sdist lacks $needed"
done
echo "$SDIST_LIST" | grep -q "__pycache__\|\.pyc$" && fail "sdist carries __pycache__"
echo "$SDIST_LIST" | grep -q "^baseline/\|^out/\|^\.venv\|^\.claude/\|^\.github/" && fail "sdist carries baseline/, out/, .venv*, .claude/ or .github/"
echo "the sdist carries the wheel's contents, tests/, samples/, schemas/, codegen/, scripts/, docs/ and the documents; nothing git-ignored"

# ---------------------------------------------------------------- 3. clean venv
step "3. a clean CPython $PYTHON_VERSION virtual environment, the wheel from dist/, dependencies from PyPI"
CLEAN_VENV="$SCRATCH/clean"
rm -rf "$CLEAN_VENV"
"$UV" venv -q --clear -p "$PYTHON_VERSION" "$CLEAN_VENV"
"$UV" pip install --python "$CLEAN_VENV/bin/python" "$WHEEL" pytest 2>&1 | tail -n 12 | sed 's/^/    /'
CLEAN="$CLEAN_VENV/bin/python"
cd "$SCRATCH"  # not the checkout: a script read from stdin has '' first on sys.path
"$CLEAN" - <<'EOF'
import importlib.metadata as m
import sys

import docx4j_py
import docx4j_xsdata
import lxml
import markdown_it

assert "site-packages" in docx4j_py.__file__, docx4j_py.__file__
print(f"    python {sys.version.split()[0]} at {sys.prefix}")
print(f"    docx4j {m.version('docx4j')} from {docx4j_py.__file__}")
print(f"    docx4j-xsdata {docx4j_xsdata.__version__} from {docx4j_xsdata.__file__}")
print(f"    lxml {m.version('lxml')}, markdown-it-py {m.version('markdown-it-py')}, pytest {m.version('pytest')}")
EOF
cd "$ROOT"

# ---------------------------------------------------------------- 4. smoke
step "4. the smoke script, from a directory outside the repository"
SMOKE="$SCRATCH/smoke"
rm -rf "$SMOKE"; mkdir -p "$SMOKE"
cp "$ROOT/samples/2010-sample1.docx" "$SMOKE/in.docx"
cat >"$SMOKE/smoke.py" <<'EOF'
"""Every packaged resource read, through the public surface, from the installed wheel."""
import os
import sys
import time

assert not os.path.exists("docx4j_py"), "the smoke must run outside the checkout"
started = time.perf_counter()
import docx4j_py  # noqa: E402

print(f"    import docx4j_py: {time.perf_counter() - started:.2f} s from {docx4j_py.__file__}")
assert "site-packages" in docx4j_py.__file__
from docx4j_py import create_package, load  # noqa: E402
from docx4j_py.docx import Document, Pt  # noqa: E402
from docx4j_py.model.content import Author  # noqa: E402
from docx4j_py.wml import el, p, r, to_xml, warm_up, wml  # noqa: E402

warm_up()  # resources/warmup_*.xml

pkg = load("in.docx")
before = len(pkg.body.paragraphs)
title = pkg.body.insert_paragraph("Smoke test", location="Start", style="Heading 1")
# Heading 1 is not in this document: the definition comes from resources/KnownStyles.xml
assert "/word/styles.xml" in pkg.last_change.parts_touched, pkg.last_change.parts_touched
pkg.body.insert_markdown("# Markdown\n\n- one\n- two\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n")
# the list creates numbering.xml from resources/openpackaging/numbering.xml
assert "/word/numbering.xml" in pkg.last_change.parts_touched
hit = pkg.find("first")[0]
pkg.author = Author("Smoke", initials="S")
pkg.change_tracking_mode = "TrackAll"
pkg.paragraph_at(hit.address).insert_text("Reviewed. ", location="Start")
pkg.find("first")[0].range(pkg.body).insert_comment("From the installed wheel.")
assert len(pkg.get_tracked_changes()) == 1
assert len(pkg.body.get_comments()) == 1
mapper = pkg.font_mapper  # resources/fonts/substitutions.json, families.json, registry.json
assert pkg.main_document_part.fonts_in_use()
outline = pkg.outline()
assert outline.stats.paragraphs > before
pkg.body.content.append(p("Typed ", r("run", bold=True)))
pkg.body.content.append(wml("<w:p><w:r><w:t>Fragment</w:t></w:r></w:p>"))
assert b"Fragment" in to_xml(el.p(content=[el.r(content=[el.t("Fragment")])])).encode()
pkg.change_tracking_mode = "Off"  # the facade below adds runs, and python-docx's runs are the direct w:r children
pkg.save("out.docx")

again = load("out.docx")
assert again.body.paragraphs[0].text == "Smoke test"
assert again.body.paragraphs[0].style == "Heading 1", again.body.paragraphs[0].style
assert again.body.get_comments()[0].content == "From the installed wheel."
assert "Reviewed." in again.body.text
assert again.skipped == [], again.skipped

doc = Document("out.docx")
doc.add_heading("Facade", level=2)
paragraph = doc.add_paragraph("Plain, ")
paragraph.add_run("bold").bold = True
paragraph.runs[1].font.size = Pt(14)
doc.add_table(rows=2, cols=2, style="Table Grid").cell(0, 0).text = "a"
doc.save("facade.docx")
assert Document("facade.docx").tables[0].cell(0, 0).text == "a"

created = create_package()  # resources/openpackaging/*.xml and the theme part
created.body.insert_paragraph("Hello World", style="Heading 1")
created.save("created.docx")
reread = load("created.docx")
assert reread.get_part("/word/theme/theme1.xml") is not None
assert reread.body.paragraphs[0].text == "Hello World"
assert reread.describe().page.orientation == "portrait"
print(f"    smoke ok: {os.getcwd()}: out.docx, facade.docx, created.docx written and read back")
EOF
(cd "$SMOKE" && "$CLEAN" smoke.py)

# ---------------------------------------------------------------- 5. the suite
step "5. the fast suite from the sdist's tests/, against the installed wheel"
SUITE="$SCRATCH/suite"
rm -rf "$SUITE"; mkdir -p "$SUITE"
tar -xzf "$SDIST" -C "$SUITE"
SUITE_ROOT="$SUITE/${NAME}-${VERSION}"
rm -rf "$SUITE_ROOT/docx4j_py"
(cd "$SUITE_ROOT" && "$CLEAN" -c 'import sys; sys.path.insert(0, "."); import docx4j_py; assert "site-packages" in docx4j_py.__file__, docx4j_py.__file__; print("    docx4j_py resolves to", docx4j_py.__file__)')
set +e
(cd "$SUITE_ROOT" && "$CLEAN" -m pytest -m "not slow" --ignore=tests/test_codegen_el.py -p no:cacheprovider >"$SCRATCH/suite.log" 2>&1)
SUITE_EXIT=$?
set -e
tail -n 3 "$SCRATCH/suite.log" | sed 's/^/    /'
[ "$SUITE_EXIT" -eq 0 ] || { echo "    (the whole log: $SCRATCH/suite.log)"; fail "the suite against the wheel, exit $SUITE_EXIT"; }

# ---------------------------------------------------------------- 6. acceptance
step "6. out/acceptance/ from the installed wheel, compared with the checkout's"
ACCEPT="$SCRATCH/accept"
rm -rf "$ACCEPT"; mkdir -p "$ACCEPT/scripts"
cp "$ROOT/scripts/acceptance.py" "$ACCEPT/scripts/"
cp -r "$ROOT/samples" "$ACCEPT/samples"
(cd "$ACCEPT" && "$CLEAN" -c 'import sys; sys.path.insert(0, "."); import docx4j_py; assert "site-packages" in docx4j_py.__file__')
(cd "$ACCEPT" && "$CLEAN" scripts/acceptance.py --out "$ACCEPT/out/acceptance" | sed 's/^/    /')
rm -rf "$ROOT/out/acceptance"
mkdir -p "$ROOT/out"
cp -r "$ACCEPT/out/acceptance" "$ROOT/out/acceptance"
echo "    copied to out/acceptance/ (the artefacts for the Word check are the wheel's)"
if [ -x "$CHECKOUT_PYTHON" ]; then
    rm -rf "$ROOT/out/acceptance-checkout"
    "$CHECKOUT_PYTHON" scripts/acceptance.py --out "$ROOT/out/acceptance-checkout" >/dev/null
    "$CLEAN" scripts/compare_acceptance.py "$ROOT/out/acceptance-checkout" "$ROOT/out/acceptance" | sed 's/^/    /'
else
    echo "    no $CHECKOUT_PYTHON: the checkout comparison skipped"
fi

echo
echo "== release check passed for $NAME $VERSION"
echo "   dist/: $(basename "$WHEEL") ($(stat -c %s "$WHEEL") bytes), $(basename "$SDIST") ($(stat -c %s "$SDIST") bytes)"
echo "   scratch: $SCRATCH"
echo "   next: open the twelve artefacts in out/acceptance/ in Word and record the check in tests/README.md"
