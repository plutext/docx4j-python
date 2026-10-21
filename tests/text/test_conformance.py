"""docx4j's text-extraction suite over this package's readers, held to the committed report.

CR-011 sections 3.6 and 4. The suite is docx4j's at ``198d015a3``, copied
into ``tests/fixtures/text-extraction/``; ``conformance.py`` is the port of
its harness and ``adapter.py`` the engine over the text model
(``docx4j_py.model.text``, CR-011 Phase B) and, for Markdown, ``to_markdown``.

- **The ratchet.** The adapter's run is compared, check by check, with the
  committed report ``tests/text/conformance.tsv``, and any difference fails: a
  check newly failing, one newly passing, a reason changed, a check added or
  gone. A change that moves a check rewrites the report in the same commit
  (``.venv-fork/bin/python tests/text/conformance.py``), so that every phase's
  progress is in the diff and nothing regresses silently. Phase A's report was
  CR-011 section 1's baseline: 941 pass, 410 fail and 233 pending of 1,584,
  3 of the 12 goldens; Phase B's is 1,146, 205 and 233, 6 of 12 (section 12);
  Phase C's is section 13's. Q7's deferred content pends by name
  (``adapter.DEFERRED``).
- **The port's gate.** With docx4j's statement set advertised and nothing
  behind it, the harness makes docx4j's 1,584 checks: 1,565 and 19 pending
  with ``html reader absent``, 1,526 and 58 with ``present``.
- **The copy**: the tables hold the rows section 3.6 counts.
- **The harness's own rules**, as docx4j's ``ExtractionConformanceTest`` and
  the javadoc of ``ExtractionConformance`` state them.
"""

from __future__ import annotations

import collections
from collections.abc import Collection, Mapping
from pathlib import Path

import conformance
import pytest
from adapter import DEFERRED, DOCX4J_IMPLEMENTED, INSPECTOR, TextsAdapter
from conformance import (
    HTML_READER_ABSENT,
    HTML_READER_PRESENT,
    LIMITATIONS,
    PHASE0,
    REPORT,
    Check,
    CheckFailed,
    Engine,
    Inspection,
    Output,
    Run,
    count,
    normalize,
    read_report,
    scenario_requirements,
    table,
    verify,
)

REWRITE = ".venv-fork/bin/python tests/text/conformance.py"


@pytest.fixture(scope="module")
def adapter_run() -> Run:
    """The suite over today's readers, run once for the module."""
    return conformance.run(TextsAdapter())


# ---------------------------------------------------------------------------
# the ratchet
# ---------------------------------------------------------------------------


def _describe(check: Check | None) -> str:
    if check is None:
        return "absent"
    return check.status + (f" ({check.reason})" if check.reason else "")


def test_the_run_is_the_committed_report(adapter_run):
    committed = {check.key: check for check in read_report(REPORT)}
    current = {check.key: check for check in adapter_run.checks}
    changed = [
        (key, committed.get(key), current.get(key))
        for key in dict.fromkeys([*committed, *current])
        if committed.get(key) != current.get(key)
    ]
    if changed:
        moves = collections.Counter(
            f"{c.status if c else 'absent'} -> {n.status if n else 'absent'}" for _key, c, n in changed
        )
        shown = "\n".join(
            f"  {' | '.join(key)}: {_describe(c)} -> {_describe(n)}" for key, c, n in changed[:25]
        )
        pytest.fail(
            f"{len(changed)} checks differ from {REPORT.name} ({dict(moves)}); {adapter_run.summary()}.\n"
            f"If the change is intended, rewrite the report with `{REWRITE}` and commit it with the change.\n"
            f"{shown}"
        )
    # the order of the rows and the bytes too, so that the file is what the run writes
    assert REPORT.read_bytes().decode("utf-8") == adapter_run.report()


def test_only_docx4js_pending_checks_the_inspectors_and_q7s_pend(adapter_run):
    """The adapter advertises docx4j's statements less the inspector's, and names Q7's checks."""
    pending = [c for c in adapter_run.checks if c.status == "PENDING"]
    reader = [c for c in pending if c.reason.endswith(HTML_READER_PRESENT)]
    deferred = [c for c in pending if c.reason.startswith("deferred (CR-011 Q7): ")]
    inspector = [c for c in pending if c not in reader and c not in deferred]
    assert len(reader) == 19
    assert {(c.fixture, c.case, c.item, c.scenario) for c in deferred} == set(DEFERRED)
    assert len(deferred) == 3 * len(DEFERRED) == 60
    assert len(inspector) == 214
    assert all(c.reason.startswith("the inspector (CR-050); ") for c in inspector)
    assert collections.Counter(c.item for c in inspector if not c.item.startswith("inspect:")) == {
        "counts": 24,
        "inspection": 4,
    }


def test_the_adapter_advertises_the_target_less_the_inspector():
    implemented = set(TextsAdapter().implemented())
    assert implemented == (DOCX4J_IMPLEMENTED - INSPECTOR) | {HTML_READER_ABSENT}
    assert HTML_READER_PRESENT not in implemented
    assert len(DOCX4J_IMPLEMENTED) == 65  # 64 statements and CR-047


def test_main_writes_the_report(tmp_path, monkeypatch, adapter_run, capsys):
    monkeypatch.setattr(conformance, "run", lambda engine, directory=PHASE0: adapter_run)
    out = tmp_path / "report.tsv"
    assert conformance.main([str(out)]) == 0
    assert out.read_bytes() == REPORT.read_bytes()
    assert adapter_run.summary() in capsys.readouterr().out


# ---------------------------------------------------------------------------
# the port's gate: docx4j's 1,584 assertions
# ---------------------------------------------------------------------------


class Claims(Engine):
    """docx4j's claims with no engine behind them: every ready check is made, and fails."""

    def __init__(self, reader: str) -> None:
        """Claims advertising docx4j's statements and one of the two reader capabilities."""
        self.reader = reader

    def implemented(self) -> Collection[str]:
        """docx4j's ``TextsAdapter.implementedStatements()`` at 198d015a3."""
        return DOCX4J_IMPLEMENTED | {self.reader}

    def extract(self, fixture: Path, view: str, scenario: str, format: str) -> Output:
        """Nothing."""
        return Output("", {})


@pytest.mark.parametrize(
    ("reader", "checked", "pending"),
    [(HTML_READER_ABSENT, 1565, 19), (HTML_READER_PRESENT, 1526, 58)],
    ids=["absent", "present"],
)
def test_the_port_makes_docx4js_checks(reader, checked, pending):
    result = conformance.run(Claims(reader))
    assert len(result.checks) == 1584
    assert (result.count_of("PASS") + result.count_of("FAIL"), result.count_of("PENDING")) == (checked, pending)


class Deferring(Claims):
    """docx4j's claims, deferring one token by name."""

    def deferred(self, fixture: str, case: str, item: str, scenario: str) -> str | None:
        """Case 1's first token in its default scenario."""
        if (case, item, scenario) == ("1", "CHOICE_ALPHA_123", "default"):
            return "deferred: a test's"
        return None


def test_a_deferred_check_pends_with_its_reason_and_nothing_else_moves():
    plain = conformance.run(Claims(HTML_READER_ABSENT)).checks
    deferring = conformance.run(Deferring(HTML_READER_ABSENT)).checks
    moved = [(a, b) for a, b in zip(plain, deferring, strict=True) if a != b]
    assert len(moved) == 3
    assert all(b.status == "PENDING" and b.reason == "deferred: a test's" for _a, b in moved)
    assert {b.view for _a, b in moved} == {"accepted", "original", "markup"}


def test_without_a_reader_docx4js_19_pend_for_one():
    result = conformance.run(Claims(HTML_READER_ABSENT))
    pending = collections.Counter((c.fixture, c.scenario) for c in result.checks if c.status == "PENDING")
    assert pending == {
        ("extraction-probe.docx", HTML_READER_PRESENT): 15,
        ("extraction-probe-inspect.docx", HTML_READER_PRESENT): 3,
        ("extraction-probe-html.docx", HTML_READER_PRESENT): 1,
    }


# ---------------------------------------------------------------------------
# the copy
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("name", "rows"),
    [
        ("expected-tokens.tsv", 285),
        ("expected-findings.tsv", 201),
        ("expected-order.tsv", 2),
        ("case-statements.tsv", 67),
        ("goldens.tsv", 12),
        ("expected-inspection.tsv", 186),
        ("expected-inspection-status.tsv", 4),
        ("expected-inspection-html.tsv", 17),
        ("expected-inspection-html-status.tsv", 1),
    ],
)
def test_the_suite_is_the_one_section_3_6_counts(name, rows):
    assert len(table(PHASE0 / name)) == rows


def test_the_suite_has_its_fixtures_and_goldens():
    findings = table(PHASE0 / "expected-findings.tsv")
    assert sum(row["channel"] == "inspect" for row in findings) == 186
    fixtures = {row["fixture"] for row in table(PHASE0 / "expected-tokens.tsv")}
    assert fixtures == {
        "extraction-probe.docx",
        "extraction-probe-external.docm",
        "extraction-probe-structure.docx",
        "extraction-probe-inspect.docx",
    }
    assert (PHASE0 / "extraction-probe-html.docx").is_file()
    for row in table(PHASE0 / "goldens.tsv"):
        assert (PHASE0 / row["golden"]).is_file(), row["golden"]
    assert len({row["case"] for row in table(PHASE0 / "case-statements.tsv")}) == 67


# ---------------------------------------------------------------------------
# the harness's rules
# ---------------------------------------------------------------------------


def test_normalisation_keeps_what_req_1201_keeps():
    # CriticMarkup's brackets, whitespace, \p{Z}, *, _, \, {, } and $ go
    assert normalize("{++A_B++} {--C--}\u00a0{>>D<<}*\\{}$\u2028") == "ABCD"
    # zero-width spaces, tag characters, both hyphens and composition stay
    kept = "\u200b\U000e0041\u2011\u00ade\u0301"
    assert normalize(kept) == kept


def test_a_token_inside_a_longer_token_of_the_suite_is_not_counted():
    text = "FLD19_RESULT FLD19_RESULT_TWO"
    assert count(text, "FLD19_RESULT") == 2
    assert count(text, "FLD19_RESULT", ["FLD19_RESULT", "FLD19_RESULT_TWO"]) == 1
    assert count(text, "FLD19_RESULT_TWO", ["FLD19_RESULT"]) == 1
    with pytest.raises(ValueError, match="Empty marker"):
        count(text, " _ ")


def test_a_marking_needs_its_total_and_one_occurrence_of_each_kind():
    verify("{--X1--}{++X1++}", "X1", "del+ins")
    with pytest.raises(CheckFailed, match="total: expected 2, got 1"):
        verify("{--X1--}", "X1", "del+ins")
    with pytest.raises(CheckFailed, match="marking del: expected 1, got 0"):
        verify("{++X1++}{++X1++}", "X1", "del+ins")
    with pytest.raises(ValueError, match="Unknown marking"):
        verify("X1", "X1", "bold")


def test_a_scenario_needs_each_of_its_options_statements():
    assert scenario_requirements("default") == []
    assert scenario_requirements("embedded: true; hidden: true") == ["REQ-512", "REQ-507"]
    assert scenario_requirements("include: comments") == ["REQ-504"]
    assert scenario_requirements(HTML_READER_ABSENT) == ["REQ-513", HTML_READER_ABSENT]
    with pytest.raises(ValueError, match="Unknown scenario"):
        scenario_requirements("colours: true")


class Raising(Claims):
    """An engine whose extraction always raises."""

    def extract(self, fixture: Path, view: str, scenario: str, format: str) -> Output:
        """A failure of the engine's."""
        raise RuntimeError(f"cannot read {fixture.name}")


def test_an_engines_exception_fails_the_checks_that_read_its_output():
    result = conformance.run(Raising(HTML_READER_ABSENT))
    failed = [c for c in result.checks if c.status == "FAIL"]
    assert result.count_of("PASS") == 0
    assert len(failed) == 1565
    reasons = collections.Counter(c.reason.partition(" ")[0] for c in failed)
    assert reasons == {"RuntimeError:": 1537, "NotImplementedError:": 28}


def _inspection_suite(directory: Path) -> list[dict[str, str]]:
    """docx4j's two-finding suite of ``inspectionComparisonIsExactByOccurrence``."""
    directory.mkdir()
    (directory / "expected-tokens.tsv").write_text(
        "case\ttoken\taccepted\toriginal\tmarkup\toption\twith_option\tfixture\tstatements\n"
    )
    (directory / "expected-findings.tsv").write_text(
        "case\tfixture\tchannel\tkind\tselector\tscenario\texpected\tstatements\n"
    )
    (directory / "goldens.tsv").write_text("fixture\tcase\tview\tscenario\tformat\tgolden\tstatements\n")
    header = (
        "fixture\tcase\tkind\tchain\tpart\tpath\tstart\tend_path\tend\taddress\tvisible\treveal\tplain\tmarkdown"
        "\tblocks\thidden\tembedded\tcomments\texcerpt_length\texcerpt_cut\tselector\tsummary\n"
    )
    one = (
        "f.docx\t1\ttext.hidden\t\t/word/document.xml\t/w:document[1]/w:body[1]/w:p[1]/w:r[1]\t0\t\t4\tbody/0\tno"
        "\toptions: hidden: true\tomitted\tomitted\tomitted\tunseen\tomitted\tomitted\t4\tfalse\tA_B\t\n"
    )
    two = one.replace("w:p[1]", "w:p[2]").replace("body/0", "body/1")
    (directory / conformance.INSPECTION_TABLE).write_text(header + one + two)
    (directory / conformance.INSPECTION_STATUS).write_text("fixture\tstatus\tskipped\tundecided\nf.docx\tCOMPLETE\t\t\n")
    return table(directory / conformance.INSPECTION_TABLE)


class Inspector(Engine):
    """An inspector that reports what it is given, and extracts nothing."""

    def __init__(self, findings: list[Mapping[str, str]]) -> None:
        """An inspector reporting `findings`."""
        self.findings = findings

    def implemented(self) -> Collection[str]:
        """The inspector's statements, without REQ-903's counts."""
        return {*(f"REQ-{n}" for n in (901, 902, 904, 905, 906, 907)), "REQ-1204", HTML_READER_ABSENT}

    def extract(self, fixture: Path, view: str, scenario: str, format: str) -> Output:
        """Never called: the suite has no token rows."""
        raise AssertionError("not called")

    def inspect(self, fixture: Path) -> Inspection:
        """The findings as given, complete, with every limitation."""
        return Inspection(self.findings, "COMPLETE", (), (), LIMITATIONS)


def test_an_inspection_is_compared_by_occurrence(tmp_path):
    """docx4j's ``inspectionComparisonIsExactByOccurrence``: missing, doubled or regrouped fails."""
    directory = tmp_path / "inspected"
    expected = _inspection_suite(directory)
    result = conformance.run(Inspector(expected), directory)
    assert "f.docx\t0\tinspection\t-\tdefault\tPASS" in result.report()
    moved = {**expected[1], "plain": "unseen"}
    for wrong in ([expected[0]], [expected[0], expected[1], expected[1]], [expected[0], moved]):
        checks = conformance.run(Inspector(wrong), directory).checks
        inspection = next(c for c in checks if c.item == "inspection")
        assert inspection.status == "FAIL", len(wrong)
        assert inspection.reason.startswith("inspection differs"), inspection.reason
