"""CR-003 Phase I, section 3.11: ``to_api_script`` as an agent uses it.

"For an agent it is reveal codes: shown a document, it learns the calls that
would make it." These are the tool-shaped sessions --- read a budgeted script,
ask why a block fell back, generate one block, run the script into a new
document --- rather than the per-rule tests, which are
``tests/content/test_api_script.py``.
"""

from __future__ import annotations

import ast
import time

import pytest
from conftest import SAMPLES, sample

from docx4j_py import create_package, load
from docx4j_py.model.content.api_script import ScriptBlock, script_blocks, to_api_script

#: Every WordprocessingML document of the corpus.
CORPUS = sorted(path.name for path in SAMPLES.glob("*.docx"))


def _run(script: str):
    """What a server would do with a script: run it into a fresh document."""
    target = create_package()
    target.id_seed = 20260916
    exec(compile(script, "<to_api_script>", "exec"), {"body": target.body})  # noqa: S102
    return target


# ---------------------------------------------------------------------------
# reveal codes
# ---------------------------------------------------------------------------


def test_reveal_codes_for_one_paragraph(report) -> None:
    """The agent asks for one block and gets the calls that make it."""
    paragraph = report.body.paragraph_at(contains="Quarterly")
    assert paragraph.to_api_script().splitlines() == [
        'p1 = body.insert_paragraph("Quarterly Report")',
        'p1.style_built_in = "Heading1"',
    ]


def test_a_budgeted_script_says_where_it_stopped(report) -> None:
    """``limit`` is the budget of section 3.4, and the tail says how to go on."""
    script = report.body.to_api_script(limit=3)
    assert script.splitlines()[-1].startswith("# ... 4 more blocks; ")
    assert "to_api_script(body, limit=None)" in script


def test_the_blocks_say_which_construct_forced_the_fallback(report) -> None:
    """``script_blocks`` is the machine-readable half: a tool result, not source."""
    report.body.insert_xml(
        '<w:p><w:hyperlink r:id="rId9"><w:r><w:t>a link</w:t></w:r></w:hyperlink></w:p>'
    )
    blocks = script_blocks(report.body)
    fell_back = [block for block in blocks if block.kind == "xml"]
    assert [block.reason for block in fell_back] == ["w:hyperlink in the paragraph"]
    assert isinstance(fell_back[0], ScriptBlock)
    assert fell_back[0].to_dict()["address"] == fell_back[0].address
    assert set(fell_back[0].to_dict()) == {"kind", "address", "exact", "reason", "lines"}


def test_the_script_rebuilds_the_document(report) -> None:
    """Run what was revealed: the same text, the same styles, the same table."""
    made = _run(report.body.to_api_script())
    assert made.body.get_text() == report.body.get_text()
    assert [paragraph.style for paragraph in made.body.paragraphs] == [
        paragraph.style for paragraph in report.body.paragraphs
    ]
    assert [table.values for table in made.body.tables] == [
        table.values for table in report.body.tables
    ]


def test_a_dry_run_does_not_leak_into_the_script(report) -> None:
    """Generating a script mutates nothing: the change log is untouched."""
    report.changes.clear()
    to_api_script(report.body)
    assert report.changes == []


# ---------------------------------------------------------------------------
# the corpus
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", CORPUS)
def test_every_corpus_script_is_python_and_deterministic(name: str) -> None:
    """``ast.parse`` on the script, and the same bytes from a second load."""
    package = load(SAMPLES / name)
    package.id_seed = 20260916
    first = to_api_script(package.body)
    ast.parse(first)
    again = load(SAMPLES / name)
    again.id_seed = 20260916
    assert to_api_script(again.body) == first


@pytest.mark.slow
@pytest.mark.parametrize("name", CORPUS)
def test_every_corpus_script_runs(name: str) -> None:
    """The whole corpus executes against a fresh body and keeps its text."""
    package = load(SAMPLES / name)
    package.id_seed = 20260916
    made = _run(to_api_script(package.body))
    assert made.body.get_text() == package.body.get_text()


@pytest.mark.slow
def test_generating_a_script_for_the_largest_sample_is_quick() -> None:
    """The budget an MCP tool call has: the biggest sample in well under a second."""
    package = sample("Symbols.docx")
    started = time.perf_counter()
    script = to_api_script(package.body)
    assert time.perf_counter() - started < 5.0
    assert len(script) > 100_000
