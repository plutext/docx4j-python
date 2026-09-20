"""The examples in the readmes execute, CR-007 section 4.5 point 5.

Every ```` ```python ```` fence of ``docs/pypi-readme.md`` (the PyPI page) and of
``README.md`` (the repository's front page) is extracted and executed **in
order, in one namespace per file**, so a later fence sees what an earlier one
bound (``pkg``, ``body``, ``document``), as a reader typing them in would. A
fence is one parametrised test, and running a single fence runs the ones
before it first, so ``-k`` selection works and a failure is reported once, on
the first fence that raised, with the later ones failing by name.

A fence in ``README.md`` that is a **fragment** --- it uses a name nothing
defines (``png_bytes``, ``package_xml``), or it demonstrates an error --- is
marked with an HTML comment on the line before it::

    <!-- fragment -->
    ```python
    picture = body.insert_inline_picture(png_bytes, ...)
    ```

and is skipped, and the skip says so. ``docs/pypi-readme.md`` may carry no
fragment at all: its fences are what PyPI shows, and every one of them runs.

**The working directory.** The fences name files as a user would --- ``in.docx``,
``report.docx``, ``in.pptx``, ``samples/invoice2013.docx`` --- so each file's
fences run with the working directory set to a temporary workspace that holds
a copy of ``samples/`` and the three files the examples open:

- ``in.docx`` and ``report.docx``: ``samples/sample-docx.docx`` with two
  paragraphs prepended, a *Chapter 1* heading and one holding the phrases the
  examples search for (``first``, ``quick brown fox``), so that a search the
  readme shows returns a hit rather than an empty list;
- ``in.pptx``: ``samples/loadAndSave.pptx``, for the generic OPC example.

Whatever a fence saves (``out.docx``, ``hello.docx``, ``edited.docx``) lands in
the workspace and is thrown away with it.

    .venv-fork/bin/python -m pytest tests/test_readme.py -q
"""

from __future__ import annotations

import os
import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

FRAGMENT_MARKER = "<!-- fragment -->"
FENCE_OPEN = re.compile(r"^```python\s*$")
FENCE_CLOSE = re.compile(r"^```\s*$")

README = ROOT / "README.md"
PYPI_README = ROOT / "docs" / "pypi-readme.md"


@dataclass(frozen=True)
class Fence:
    path: Path
    line: int  # 1-based line of the opening ```python
    code: str
    fragment: bool

    @property
    def id(self) -> str:
        return f"{self.path.name}:{self.line}"


def extract_fences(path: Path) -> list[Fence]:
    """Every ```` ```python ```` fence of ``path``, in order, with its marker."""
    fences: list[Fence] = []
    lines = path.read_text(encoding="utf-8").splitlines()
    previous_nonblank = ""
    i = 0
    while i < len(lines):
        line = lines[i]
        if FENCE_OPEN.match(line):
            start = i + 1
            body: list[str] = []
            i += 1
            while i < len(lines) and not FENCE_CLOSE.match(lines[i]):
                body.append(lines[i])
                i += 1
            fences.append(
                Fence(
                    path=path,
                    line=start,
                    code="\n".join(body) + "\n",
                    fragment=previous_nonblank.strip() == FRAGMENT_MARKER,
                )
            )
            previous_nonblank = ""
        elif line.strip():
            previous_nonblank = line
        i += 1
    return fences


# ---------------------------------------------------------------------------
# the workspace and the runner
# ---------------------------------------------------------------------------


def make_workspace(root: Path) -> Path:
    """A directory with ``samples/``, ``in.docx``, ``report.docx`` and ``in.pptx``."""
    from docx4j_py import load

    shutil.copytree(ROOT / "samples", root / "samples")
    shutil.copy(ROOT / "samples" / "loadAndSave.pptx", root / "in.pptx")

    pkg = load(root / "samples" / "sample-docx.docx")
    pkg.body.insert_paragraph(
        "My first document: the quick brown fox jumps over the lazy dog.",
        location="Start",
    )
    pkg.body.insert_paragraph("Chapter 1", location="Start", style="Heading 1")
    pkg.save(root / "in.docx")
    shutil.copy(root / "in.docx", root / "report.docx")
    return root


class FenceRunner:
    """Executes a file's fences in order, once each, in one namespace."""

    def __init__(self, fences: list[Fence], workspace: Path) -> None:
        self.fences = fences
        self.workspace = workspace
        self.namespace: dict[str, object] = {"__name__": "__readme__"}
        self.done = 0
        self.failed: tuple[Fence, BaseException] | None = None

    def run_through(self, index: int) -> None:
        """Run every fence up to and including ``index`` that has not run yet."""
        while self.done <= index:
            fence = self.fences[self.done]
            if self.failed is not None:
                first, error = self.failed
                pytest.fail(
                    f"{fence.id} not run: {first.id} raised {type(error).__name__}: {error}",
                    pytrace=False,
                )
            self.done += 1
            if fence.fragment:
                continue
            cwd = os.getcwd()
            os.chdir(self.workspace)
            try:
                code = compile(fence.code, fence.id, "exec")
                exec(code, self.namespace)  # noqa: S102 - that is the test
            except BaseException as error:
                self.failed = (fence, error)
                raise
            finally:
                os.chdir(cwd)


def _runner(request: pytest.FixtureRequest, path: Path) -> FenceRunner:
    tmp = request.getfixturevalue("tmp_path_factory").mktemp(path.stem)
    return FenceRunner(extract_fences(path), make_workspace(tmp))


@pytest.fixture(scope="module")
def readme_runner(request: pytest.FixtureRequest) -> FenceRunner:
    return _runner(request, README)


@pytest.fixture(scope="module")
def pypi_runner(request: pytest.FixtureRequest) -> FenceRunner:
    return _runner(request, PYPI_README)


README_FENCES = extract_fences(README)
PYPI_FENCES = extract_fences(PYPI_README)


# ---------------------------------------------------------------------------
# the tests
# ---------------------------------------------------------------------------


def test_both_readmes_have_python_fences() -> None:
    assert len(PYPI_FENCES) >= 5
    assert len(README_FENCES) >= 30


def test_the_pypi_readme_carries_no_fragment() -> None:
    """What PyPI shows runs, all of it: no fence there may be marked."""
    assert [f.id for f in PYPI_FENCES if f.fragment] == []
    assert FRAGMENT_MARKER not in PYPI_README.read_text(encoding="utf-8")


def test_the_readme_documents_the_marker() -> None:
    """The marker is explained where a reader adding a fence will see it."""
    text = README.read_text(encoding="utf-8")
    running_it = text.index("## Running it")
    assert FRAGMENT_MARKER in text[running_it:], "README.md's 'Running it' must explain the marker"


@pytest.mark.parametrize("index", range(len(PYPI_FENCES)), ids=[f.id for f in PYPI_FENCES])
def test_pypi_readme_fence(pypi_runner: FenceRunner, index: int) -> None:
    pypi_runner.run_through(index)


@pytest.mark.parametrize("index", range(len(README_FENCES)), ids=[f.id for f in README_FENCES])
def test_readme_fence(readme_runner: FenceRunner, index: int) -> None:
    fence = README_FENCES[index]
    if fence.fragment:
        readme_runner.run_through(index)
        pytest.skip(f"{fence.id} is marked {FRAGMENT_MARKER}")
    readme_runner.run_through(index)
