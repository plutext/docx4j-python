# Releasing

The distribution `docx4j` (import name `docx4j_py`) is published to PyPI from GitHub Actions
(`.github/workflows/publish.yml`), the way docx4j-core-ts publishes `@docx4j/core-ts` to npm and
the xsdata fork publishes `docx4j-xsdata`. The design of the release is
[CR-007](docs/change-requests/CR-007-pypi-release.md); this file is the procedure.

## Publishing target

Publishing uses PyPI **trusted publishing**: no API token is stored anywhere. The trusted
publishers on pypi.org and test.pypi.org name the GitHub repository allowed to publish: owner
`plutext`, repository `docx4j-python`, workflow `publish.yml`, no environment. The first release
uses a **pending publisher** registered under the same settings before the project existed, so
the first upload goes through the workflow like every later one. Renaming the workflow file breaks
publishing until both settings are changed to match.

`publish.yml` runs on a GitHub **release** (`types: [published]`) and on `workflow_dispatch`. The
release path checks that the release tag is `v` plus `pyproject.toml`'s version, builds the sdist
and the wheel with `python -m build`, runs `twine check --strict`, and uploads to PyPI through
OIDC. The dispatch path is the same build uploaded to **TestPyPI**, the dry run that shows the
name, the rendered readme and the wheel before the version is taken on PyPI.

A version once published is never reused, on PyPI or on TestPyPI, even after a deletion: a
problem found after publishing ships as the next patch version. TestPyPI has the same rule, so a
second dry run of the same version fails at the upload; that is expected.

## Dependencies

`pyproject.toml` depends on `docx4j-xsdata[lxml]>=26.2.1,<26.3`, the xsdata fork every generated
module imports, and on `markdown-it-py>=3.0`. Consumers resolve those ranges from PyPI; CI
(`test.yml`) installs the fork **from PyPI** too, so what CI tests is what a consumer gets. The
checkout's `.venv-fork` has the fork installed editable from `~/git/docx4j-xsdata` and is for
development; nothing about a release is proved by it.

**This package is never published against an unreleased fork.** A release that needs a change in
the fork waits for the fork's release (`~/git/docx4j-xsdata/docs/fork/REBASING.md`, "Releasing"),
then raises the bound here: a fork release on the same upstream base (`26.2.2`) is additive by the
fork's rule and needs only the lower bound raised (`>=26.2.2,<26.3`) when this package relies on
what it added; a fork on a new upstream base (`26.3.1`) means the bindings are regenerated
(`codegen/generate.sh --check` against the new fork) and both bounds move (`>=26.3.1,<26.4`).
Change the pin in `dependencies` and in the `dev` extra together, run the gates, and commit.

## Gates

Every gate is a script, so that a release is the script and the Word check, not a memory of what
was done last time.

```bash
.venv-fork/bin/python -m pytest                   # the full suite on the checkout
codegen/generate.sh --check                        # the bindings regenerate byte identical
scripts/release_check.sh                           # the wheel, not the checkout (below)
```

`scripts/release_check.sh` builds the sdist and the wheel in a scratch environment, runs `twine
check --strict`, lists the wheel and asserts what CR-007 section 4.7 says ships (`resources/`,
`py.typed`, `LICENSE`, `NOTICE`, no `__pycache__`), makes a **clean Python 3.12 virtual
environment** with `uv`, installs the wheel from `dist/` with its dependencies resolved **from
PyPI**, runs a smoke script from a directory outside the repository (load a corpus document, edit
through the content API and the facade, save, reload, `warm_up()`), runs the fast suite from the
**sdist's** `tests/` against the installed wheel (the sdist unpacked with its `docx4j_py/` removed,
so the import resolves to site-packages), and regenerates `out/acceptance/` from the installed
wheel. Then the manual gate: open the twelve artefacts in Word, as `tests/README.md` describes, and
record the check there. Packaging is a marshalling input --- a resource missing from the wheel
changes what is written --- so the Word check is not waived for a release.

The remaining gate is `tests/test_readme.py`, part of the suite: every ```` ```python ```` fence of
`docs/pypi-readme.md`, the PyPI page, executes.

## Steps

```bash
# 1. If the release needs a newer docx4j-xsdata, release that first and raise the bound here
#    (see "Dependencies").

# 2. Set the version in pyproject.toml, the next one after the latest on PyPI
#    (pip index versions docx4j), and write its CHANGELOG.md entry: one paragraph per change
#    request, linking the CR's implementation notes. No tag yet.

# 3. The gates
.venv-fork/bin/python -m pytest && codegen/generate.sh --check && scripts/release_check.sh
#    then the Word check of out/acceptance/, recorded in tests/README.md

# 4. Commit and push main; CI (test.yml) must pass
git commit -am "Version 0.1.0"
git push origin main

# 5. Dry run on TestPyPI: Actions tab, publish.yml, "Run workflow" on main. Then check
#    https://test.pypi.org/project/docx4j/ --- the version, the readme rendered, the wheel and
#    the sdist listed, the URLs and the classifiers.

# 6. Tag and push the tag (the workflow checks the tag is v plus the version)
git tag -a v0.1.0 -m "docx4j 0.1.0"
git push origin v0.1.0
```

7. Publish: create a GitHub release for the tag at
   <https://github.com/plutext/docx4j-python/releases/new> (choose the existing tag, title
   `v0.1.0`, notes from the `CHANGELOG.md` entry, "Set as a pre-release" unticked, then **Publish
   release**; a saved draft does not publish).

Publishing the release runs `publish.yml`: it checks out the tag, fails unless the tag is `v` plus
`pyproject.toml`'s version, builds, runs `twine check`, and uploads to PyPI. Then confirm
<https://pypi.org/project/docx4j/> resolves and `pip install docx4j==0.1.0` into a clean
environment imports.

If the workflow fails before the upload, fix the problem on `main`, delete the release and the
tag (the version was not published, so the number is still free), and start again from step 4's
commit. If only the upload fails (a publisher setting), fix the setting and **Re-run failed
jobs**; the tag and the release stay.

## What ships

The wheel (`docx4j-<version>-py3-none-any.whl`, pure Python, about 1 MB compressed, 6 MB installed):
`docx4j_py/` --- the 238 modules, `resources/` (the warm-up parts, `KnownStyles.xml`, docx4j's
default styles, numbering and font table, the three theme parts and the font tables), `py.typed`
--- and `LICENSE` and `NOTICE` under `dist-info/licenses/`. Nothing from `codegen/`, `scripts/` or
`baseline/`, and no `__pycache__`.

The sdist (`docx4j-<version>.tar.gz`): the wheel's contents plus `tests/` with its fixtures,
`samples/`, `schemas/`, `codegen/` with `.xsdata.phase-b.xml`, `scripts/`, `docs/`,
`pyproject.toml`, the READMEs, `CHANGELOG.md`, `RELEASING.md`, `REPORT.md`, `LICENSE` and
`NOTICE`, so that `pip download --no-binary :all: docx4j` gets something whose suite runs. Not
`baseline/`, `out/`, `.venv*`, `.claude/` or `.github/`. `scripts/release_check.sh` prints both
listings.
