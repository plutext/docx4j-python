"""docx4j's text-extraction suites over this package's readers, held to the committed reports.

CR-011 sections 3.6 and 4; CR-027 section 4.4. The suites are docx4j's at its
release tag ``docx4j-17.3.5`` (``64fae4cd0``), copied into
``tests/fixtures/text-extraction/`` (``SHA1SUMS``): ``phase0/`` and, beside it,
docx4j CR-069's ``dpp/``. ``conformance.py`` is the port of docx4j's harness and
``adapter.py`` the engine over the text model (``docx4j_py.model.text``).

- **The ratchet.** Each suite's run is compared, check by check, with its
  committed report (``tests/text/conformance.tsv``, ``conformance-dpp.tsv``),
  and any difference fails: a check newly failing, one newly passing, a reason
  changed, a check added or gone. A change that moves a check rewrites the
  reports in the same commit (``.venv-fork/bin/python
  tests/text/conformance.py``), so that every phase's progress is in the diff
  and nothing regresses silently. CR-011's phases took ``phase0/`` at
  ``198d015a3`` from 941 pass, 410 fail and 233 pending of 1,584 to 1,291, 0
  and 293; CR-027 Phase A re-pinned it (section 10). Q7's deferred content
  pends by name (``adapter.DEFERRED``), and what docx4j advertises that this
  package does not build yet with its reason (``adapter.NOT_ADVERTISED``).
- **The port's gate.** With docx4j's statement set at the tag advertised and
  nothing behind it, the harness makes docx4j's checks: ``phase0/``'s 1,772,
  36 pending with ``html reader absent`` and 40 with ``present``; ``dpp/``'s
  192, none pending without a reader and 2 with one. Under ``absent`` these are
  docx4j's own run's checks, in its order, with its pending rows and reasons.
- **The copy**: every file ``SHA1SUMS`` lists, each docx4j's blob at the tag
  where the sibling checkout has it; the tables hold the rows CR-027 counts.
- **The harness's own rules**, as docx4j's ``ExtractionConformanceTest`` and
  the javadoc of ``ExtractionConformance`` state them.
"""

from __future__ import annotations

import collections
import hashlib
import subprocess
from collections.abc import Collection, Mapping
from pathlib import Path

import conformance
import pytest
from adapter import BEYOND_DOCX4J, DEFERRED, DOCX4J_IMPLEMENTED, INSPECTOR, NOT_ADVERTISED, TextsAdapter
from conformance import (
    DPP,
    DPP_REPORT,
    HTML_READER_ABSENT,
    HTML_READER_PRESENT,
    LIMITATIONS,
    PHASE0,
    REPORT,
    SUITE,
    Check,
    CheckFailed,
    Engine,
    Inspection,
    Output,
    Run,
    count,
    entries,
    normalize,
    pending_reason,
    read_report,
    scenario_requirements,
    table,
    verify,
    verify_word_save,
    word_save_lines,
)

REWRITE = ".venv-fork/bin/python tests/text/conformance.py"
#: The sibling checkout, for the copy's provenance (skipped without it).
DOCX4J = conformance.ROOT.parent / "docx4j"
TAG = "docx4j-17.3.5"
SOURCE = "docx4j-core-tests/src/test/resources/text-extraction"


@pytest.fixture(scope="module")
def adapter_run() -> Run:
    """``phase0/`` over today's readers, run once for the module."""
    return conformance.run(TextsAdapter())


@pytest.fixture(scope="module")
def dpp_run() -> Run:
    """``dpp/`` over today's readers, run once for the module."""
    return conformance.run(TextsAdapter(), DPP)


# ---------------------------------------------------------------------------
# the ratchet
# ---------------------------------------------------------------------------


def _describe(check: Check | None) -> str:
    if check is None:
        return "absent"
    return check.status + (f" ({check.reason})" if check.reason else "")


@pytest.mark.parametrize(("suite", "report"), [("adapter_run", REPORT), ("dpp_run", DPP_REPORT)], ids=["phase0", "dpp"])
def test_the_run_is_the_committed_report(suite, report, request):
    adapter_run = request.getfixturevalue(suite)
    committed = {check.key: check for check in read_report(report)}
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
            f"{len(changed)} checks differ from {report.name} ({dict(moves)}); {adapter_run.summary()}.\n"
            f"If the change is intended, rewrite the reports with `{REWRITE}` and commit them with the change.\n"
            f"{shown}"
        )
    # the order of the rows and the bytes too, so that the file is what the run writes
    assert report.read_bytes().decode("utf-8") == adapter_run.report()


def _pending_by_cause(checks: list[Check]) -> dict[str, list[Check]]:
    """The pending checks by why they pend: an HTML reader, the other environment, Q7, the inspector, CR-069."""
    out: dict[str, list[Check]] = collections.defaultdict(list)
    for c in checks:
        if c.status != "PENDING":
            continue
        if c.reason.startswith("deferred (CR-011 Q7): "):
            out["deferred"].append(c)
        elif c.reason.endswith("not this run's environment"):
            out["environment"].append(c)
        elif c.reason.endswith(HTML_READER_PRESENT):
            out["reader"].append(c)
        elif c.reason.startswith("the inspector (CR-050); "):
            out["inspector"].append(c)
        else:
            out["not advertised"].append(c)
    return out


def test_only_docx4js_pending_checks_the_inspectors_and_q7s_pend(adapter_run):
    """``phase0/``: docx4j's 36, the inspector's, and Q7's checks pend; nothing for what is not advertised."""
    pending = _pending_by_cause(adapter_run.checks)
    assert {cause: len(checks) for cause, checks in pending.items()} == {
        "reader": 33,  # docx4j's 15 extraction checks under html reader present, and 18 inspect rows of that environment
        "environment": 3,  # the other environment's rows, by count, for three fixtures
        "deferred": 60,
        "inspector": 382,
    }
    assert {(c.fixture, c.case, c.item, c.scenario) for c in pending["deferred"]} == set(DEFERRED)
    assert len(pending["deferred"]) == 3 * len(DEFERRED) == 60
    inspector = [*pending["inspector"], *(c for c in pending["reader"] if c.reason.startswith("the inspector"))]
    assert len(inspector) == 400
    whole = collections.Counter(c.item for c in inspector if c.item in ("counts", "inspection"))
    assert whole == {"counts": 30, "inspection": 5}  # five fixtures' inspections, each under six outputs
    # docx4j's own pending checks, without an HTML reader, are pending here too
    docx4j = conformance.run(Claims(HTML_READER_ABSENT))
    ours = {c.key for c in adapter_run.checks if c.status == "PENDING"}
    assert {c.key for c in docx4j.checks if c.status == "PENDING"} <= ours


def test_dpp_holds_what_holds_before_cr_069(dpp_run):
    """``dpp/``: the stored ``H₂O`` checked in three views; the rest pends on CR-069's phase or on REQ-1002.

    docx4j's ``dppSuiteIsCheckedOrPendingOnItsPhase``, as this package
    advertises: its CR-069 rows (REQ-607, REQ-521, REQ-710, the inspector's
    ``text.formatting_join``) come with CR-029, ``profile: word-text`` with
    CR-027 Phase D, so Word's 22 saves pend with them.
    """
    report = dpp_run.report()
    for view in ("accepted", "original", "markup"):
        assert f"caps/caps-lang5.docx\t5\tH\u2082O\t{view}\tdefault\tPASS\t" in report
    assert (dpp_run.count_of("PASS"), dpp_run.count_of("FAIL"), dpp_run.count_of("PENDING")) == (3, 0, 189)
    reasons = collections.Counter(c.reason.partition(";")[0] for c in dpp_run.checks if c.status == "PENDING")
    assert reasons == {
        "CR-069 phase 1": 118,
        "CR-069 phase 2": 7,
        "CR-069 phase 3": 52,
        "CR-069 phase 4": 4,
        "the inspector (CR-050)": 2,
        "phase 1": 5,
        "phase 2": 1,
    }
    saves = [c for c in dpp_run.checks if c.item.startswith("word save:")]
    assert len(saves) == 22 and all(c.status == "PENDING" and "REQ-1002" in c.reason for c in saves)
    assert {c.scenario for c in saves} == {"profile: word-text"}
    # the six Phase D checks: REQ-1002 alone missing
    assert sum(c.reason.endswith("; REQ-1002") for c in saves) == 6


def test_the_adapter_advertises_the_target_less_the_inspector():
    implemented = set(TextsAdapter().implemented())
    want = (DOCX4J_IMPLEMENTED - INSPECTOR - set(NOT_ADVERTISED)) | set(BEYOND_DOCX4J) | {HTML_READER_ABSENT}
    assert implemented == want
    assert HTML_READER_PRESENT not in implemented
    assert len(DOCX4J_IMPLEMENTED) == 70  # 68 statements, CR-047 and text.formatting_join
    assert set(NOT_ADVERTISED) == {"REQ-1002", "REQ-607", "REQ-521", "REQ-710", "text.formatting_join"}
    assert set(NOT_ADVERTISED) <= DOCX4J_IMPLEMENTED
    # CR-032: REQ-206 beyond docx4j's tag, which no check of the suites at the tag needs
    assert set(BEYOND_DOCX4J) == {"REQ-206"} and not set(BEYOND_DOCX4J) & DOCX4J_IMPLEMENTED


def test_main_writes_the_reports(tmp_path, monkeypatch, adapter_run, dpp_run, capsys):
    runs = {PHASE0: adapter_run, DPP: dpp_run}
    monkeypatch.setattr(conformance, "run", lambda engine, directory=PHASE0: runs[directory])
    out, dpp = tmp_path / "report.tsv", tmp_path / "dpp.tsv"
    assert conformance.main([str(out), str(dpp)]) == 0
    assert out.read_bytes() == REPORT.read_bytes()
    assert dpp.read_bytes() == DPP_REPORT.read_bytes()
    printed = capsys.readouterr().out
    assert adapter_run.summary() in printed and dpp_run.summary() in printed


# ---------------------------------------------------------------------------
# the port's gate: docx4j's checks
# ---------------------------------------------------------------------------


class Claims(Engine):
    """docx4j's claims with no engine behind them: every ready check is made, and fails."""

    def __init__(self, reader: str) -> None:
        """Claims advertising docx4j's statements and one of the two reader capabilities."""
        self.reader = reader

    def implemented(self) -> Collection[str]:
        """docx4j's ``TextsAdapter.implementedStatements()`` at the tag."""
        return DOCX4J_IMPLEMENTED | {self.reader}

    def extract(self, fixture: Path, view: str, scenario: str, format: str) -> Output:
        """Nothing."""
        return Output("", {})


@pytest.mark.parametrize(
    ("directory", "reader", "checks", "pending"),
    [
        (PHASE0, HTML_READER_ABSENT, 1772, 36),
        (PHASE0, HTML_READER_PRESENT, 1772, 40),
        (DPP, HTML_READER_ABSENT, 192, 0),
        (DPP, HTML_READER_PRESENT, 192, 2),
    ],
    ids=["phase0-absent", "phase0-present", "dpp-absent", "dpp-present"],
)
def test_the_port_makes_docx4js_checks(directory, reader, checks, pending):
    result = conformance.run(Claims(reader), directory)
    assert (len(result.checks), result.count_of("PENDING")) == (checks, pending)


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


def test_without_a_reader_docx4js_36_pend_for_one():
    result = conformance.run(Claims(HTML_READER_ABSENT))
    pending = collections.Counter((c.fixture, c.scenario) for c in result.checks if c.status == "PENDING")
    assert pending == {
        ("extraction-probe.docx", HTML_READER_PRESENT): 15,
        ("extraction-probe-inspect.docx", HTML_READER_PRESENT): 5,
        ("extraction-probe-html.docx", HTML_READER_PRESENT): 16,
    }
    # each fixture with rows of the other environment pends them by count, and is compared on the rest
    report = result.report()
    for fixture, n in (("extraction-probe.docx", "2 rows"), ("extraction-probe-inspect.docx", "1 row")):
        assert f"{fixture}\t0\tinspection\t-\t{HTML_READER_PRESENT}\tPENDING\t{n} of {HTML_READER_PRESENT}," in report
    assert f"extraction-probe-html.docx\t0\tinspection\t-\t{HTML_READER_ABSENT}\tFAIL" in report


# ---------------------------------------------------------------------------
# the copy
# ---------------------------------------------------------------------------


def _sums() -> dict[str, str]:
    return dict(reversed(line.split("  ", 1)) for line in (SUITE / "SHA1SUMS").read_text(encoding="utf-8").splitlines())


def test_the_copy_is_every_file_sha1sums_lists():
    """Byte for byte, and nothing else beside it but the two files copied from elsewhere in docx4j."""
    listed = _sums()
    for name, digest in listed.items():
        assert hashlib.sha1((SUITE / name).read_bytes()).hexdigest() == digest, name
    present = sorted(str(p.relative_to(SUITE)) for p in SUITE.rglob("*") if p.is_file())
    others = {
        "SHA1SUMS",
        "understood-namespaces.txt",
        *(f"level-owner/{p.name}" for p in (SUITE / "level-owner").iterdir()),
    }
    assert present == sorted([*listed, *others])
    assert len(listed) == 84
    assert not (PHASE0 / "expected-inspection-html-status.tsv").exists()  # 4a37702eb deleted it


@pytest.mark.skipif(not (DOCX4J / ".git").exists(), reason="needs the sibling checkout ../docx4j")
def test_each_file_is_docx4js_at_the_tag():
    for name, digest in _sums().items():
        blob = subprocess.run(
            ["git", "-C", str(DOCX4J), "show", f"{TAG}:{SOURCE}/{name}"], capture_output=True, check=True
        ).stdout
        assert hashlib.sha1(blob).hexdigest() == digest, name


@pytest.mark.parametrize(
    ("directory", "name", "rows"),
    [
        (PHASE0, "expected-tokens.tsv", 319),
        (PHASE0, "expected-findings.tsv", 257),
        (PHASE0, "expected-order.tsv", 2),
        (PHASE0, "case-statements.tsv", 82),
        (PHASE0, "goldens.tsv", 12),
        (PHASE0, "expected-inspection.tsv", 242),
        (PHASE0, "expected-inspection-status.tsv", 8),
        (PHASE0, "expected-inspection-html.tsv", 17),
        (DPP, "expected-tokens.tsv", 34),
        (DPP, "expected-findings.tsv", 8),
        (DPP, "goldens.tsv", 2),
        (DPP, "word-saves.tsv", 22),
    ],
)
def test_the_suite_is_the_one_cr_027_counts(directory, name, rows):
    assert len(table(directory / name)) == rows


def test_the_suite_has_its_fixtures_and_goldens():
    findings = table(PHASE0 / "expected-findings.tsv")
    assert sum(row["channel"] == "inspect" for row in findings) == 242
    tokens = table(PHASE0 / "expected-tokens.tsv")
    assert {row["fixture"] for row in tokens} == {
        "extraction-probe.docx",
        "extraction-probe-external.docm",
        "extraction-probe-structure.docx",
        "extraction-probe-inspect.docx",
    }
    assert (PHASE0 / "extraction-probe-html.docx").is_file()
    for directory in (PHASE0, DPP):
        for row in table(directory / "goldens.tsv"):
            assert (directory / row["golden"]).is_file(), row["golden"]
    for row in table(DPP / "word-saves.tsv"):
        assert (DPP / row["fixture"]).is_file() and (DPP / row["save"]).is_file(), row["save"]
    # the case statements: 82 cases, each token row citing its case's statements and fixture
    cases = {row["case"]: row for row in table(PHASE0 / "case-statements.tsv")}
    assert sorted(cases, key=int) == [str(n) for n in range(1, 83)]
    for row in tokens:
        assert (row["statements"], row["fixture"]) == (cases[row["case"]]["statements"], cases[row["case"]]["fixture"])
    assert cases["33"]["statements"] == "REQ-514, REQ-902"


def test_the_builders_packages_carry_its_timestamp_and_words_saves_need_not():
    """``ExtractionFixtureChecks.entries(..., built)``: the builder's fixed 1980 timestamp, or a package as it is."""
    for fixture in ("extraction-probe.docx", "extraction-probe-external.docm", "extraction-probe-structure.docx",
                    "extraction-probe-inspect.docx", "extraction-probe-html.docx"):  # fmt: skip
        assert "word/document.xml" in entries((PHASE0 / fixture).read_bytes())
    made = (DPP / "bisect" / "C02.docx").read_bytes()
    with pytest.raises(ValueError, match="ZIP timestamp"):
        entries(made)
    assert "word/document.xml" in entries(made, built=False)
    for row in table(DPP / "word-saves.tsv"):
        assert "word/document.xml" in entries((DPP / row["fixture"]).read_bytes(), built=False)


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
    # docx4j CR-069's three (993e385e7)
    assert scenario_requirements("drop_caps: asis") == ["REQ-521"]
    assert scenario_requirements("small_caps: upper") == ["REQ-607"]
    assert scenario_requirements("profile: word-text") == ["REQ-1002"]
    with pytest.raises(ValueError, match="Unknown scenario"):
        scenario_requirements("colours: true")


def test_a_pending_check_names_cr_069s_phase_else_cr_046s_or_the_inspector():
    """docx4j's ``pendingReason``: CR-069's latest phase among the missing statements first."""
    assert pending_reason(["REQ-1002", "REQ-607"], ["REQ-1002", "REQ-607"]) == "CR-069 phase 1"
    assert pending_reason(["REQ-1203", "REQ-521", "REQ-607", "REQ-710"], ["REQ-521", "REQ-607", "REQ-710"]) == (
        "CR-069 phase 3"
    )
    inspector = ["REQ-901", "REQ-902", "text.formatting_join"]
    assert pending_reason(inspector, inspector) == "CR-069 phase 4"
    assert pending_reason(["REQ-901", "REQ-902"], ["REQ-901"]) == "the inspector (CR-050)"
    assert pending_reason(["REQ-1002", "REQ-605"], ["REQ-1002"]) == "phase 2"
    assert pending_reason(["REQ-1002"], ["REQ-1002"]) == "phase 1"


def test_a_word_save_is_read_by_its_encoding_and_compared_by_line(tmp_path):
    """docx4j's ``wordSaveIsReadByItsEncodingAndComparedByLine``: REQ-1002's NOTE, made exact."""
    cp1252 = tmp_path / "cp1252.txt"
    cp1252.write_bytes("L1_ H?O caf\u00e9\r\nC04_\r1.3161.335\rL3_ end.\r\n".encode("windows-1252"))
    text = "L1_ H\u2082O caf\u00e9\nC04_\t1.3161.335\nL3_ end.\n"
    verify_word_save(cp1252, "windows-1252", text)
    # a character Word did not write; a break between cells lost
    for wrong in (text.replace("H\u2082O", "H2O"), text.replace("C04_\t", "C04_ ")):
        with pytest.raises(CheckFailed, match=r"line \d+: Word wrote"):
            verify_word_save(cp1252, "windows-1252", wrong)
    with pytest.raises(CheckFailed, match="no byte order mark"):
        verify_word_save(cp1252, "UTF-8", text)
    utf8 = tmp_path / "utf8.txt"
    utf8.write_bytes("\ufeff\u0130STANBUL F\u0130LM\r\n".encode())
    verify_word_save(utf8, "UTF-8", "\u0130STANBUL F\u0130LM\n")
    with pytest.raises(CheckFailed, match="line 1: Word wrote"):
        verify_word_save(utf8, "UTF-8", "ISTANBUL FILM\n")
    with pytest.raises(CheckFailed, match="a byte order mark"):
        verify_word_save(utf8, "windows-1252", "\u0130STANBUL F\u0130LM\n")
    # the lines: CR LF, CR, LF and a tab end one; trailing empty lines dropped
    assert word_save_lines("a\r\nb\rc\nd\te\r\n\r\n\n") == ["a", "b", "c", "d", "e"]
    assert word_save_lines("a\n\nb\n") == ["a", "", "b"]


@pytest.mark.parametrize("row", table(DPP / "word-saves.tsv"), ids=lambda row: row["save"])
def test_words_saves_are_their_tables_encoding(row):
    """Every save of ``word-saves.tsv`` decodes as its row says: a mark exactly where the row says UTF-8."""
    data = (DPP / row["save"]).read_bytes()
    assert data.startswith(b"\xef\xbb\xbf") == (row["encoding"] == "UTF-8")
    assert word_save_lines(data.decode("utf-8-sig" if row["encoding"] == "UTF-8" else "windows-1252"))


class Raising(Claims):
    """An engine whose extraction always raises."""

    def extract(self, fixture: Path, view: str, scenario: str, format: str) -> Output:
        """A failure of the engine's."""
        raise RuntimeError(f"cannot read {fixture.name}")


def test_an_engines_exception_fails_the_checks_that_read_its_output():
    result = conformance.run(Raising(HTML_READER_ABSENT))
    failed = [c for c in result.checks if c.status == "FAIL"]
    assert result.count_of("PASS") == 0
    assert len(failed) == 1736
    reasons = collections.Counter(c.reason.partition(" ")[0] for c in failed)
    # the extractions', and the inspector's: five inspections and their counts under six outputs
    assert reasons == {"RuntimeError:": 1701, "NotImplementedError:": 35}


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

    def __init__(self, findings: list[Mapping[str, str]], req_513: bool = True) -> None:
        """An inspector reporting `findings`, claiming REQ-513 (an altChunk read or skipped) unless told not to."""
        self.findings = findings
        self.req_513 = req_513

    def implemented(self) -> Collection[str]:
        """The inspector's statements, without REQ-903's counts."""
        claimed = {*(f"REQ-{n}" for n in (901, 902, 904, 905, 906, 907)), "REQ-1204", HTML_READER_ABSENT}
        return claimed | {"REQ-513"} if self.req_513 else claimed

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
    # since CR-050 phase 4 the inspection's scenario is the engine's environment
    assert f"f.docx\t0\tinspection\t-\t{HTML_READER_ABSENT}\tPASS" in result.report()
    moved = {**expected[1], "plain": "unseen"}
    for wrong in ([expected[0]], [expected[0], expected[1], expected[1]], [expected[0], moved]):
        checks = conformance.run(Inspector(wrong), directory).checks
        inspection = next(c for c in checks if c.item == "inspection")
        assert inspection.status == "FAIL", len(wrong)
        assert inspection.reason.startswith("inspection differs"), inspection.reason


def test_an_inspection_is_compared_in_the_engines_environment(tmp_path):
    """CR-050 phase 4: the rows of ``any`` and of the environment, with its status; the other's pend by count."""
    directory = tmp_path / "capabilities"
    rows = _inspection_suite(directory)
    header, one, two = (directory / conformance.INSPECTION_TABLE).read_text().splitlines()
    present = two.replace("A_B", "C_D")
    (directory / conformance.INSPECTION_TABLE).write_text(
        f"{header}\tcapability\n{one}\t{conformance.ANY}\n{two}\t{HTML_READER_ABSENT}\n{present}\t{HTML_READER_PRESENT}\n"
    )
    (directory / conformance.INSPECTION_STATUS).write_text(
        "fixture\tcapability\tstatus\tskipped\tundecided\n"
        f"f.docx\t{HTML_READER_ABSENT}\tCOMPLETE\t\t\nf.docx\t{HTML_READER_PRESENT}\tPARTIAL\tx@|/word/a.htm|\t\n"
    )
    checks = conformance.run(Inspector(rows), directory).checks
    inspection = [c for c in checks if c.item == "inspection"]
    assert [(c.scenario, c.status, c.reason) for c in inspection] == [
        (HTML_READER_PRESENT, "PENDING", f"1 row of {HTML_READER_PRESENT}, not this run's environment"),
        (HTML_READER_ABSENT, "PASS", ""),
    ]
    # a fixture with rows of an environment needs REQ-513, which this inspector does not claim
    statuses = [c for c in conformance.run(Inspector(rows, req_513=False), directory).checks if c.item == "inspection"]
    assert statuses[1].status == "PENDING" and statuses[1].reason.endswith("REQ-513")
