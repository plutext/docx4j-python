"""docx4j's text-extraction conformance harness, ported. CR-011 sections 3.6 and 4.

Not a test: the port of ``org.docx4j.textextraction.ExtractionConformance`` at
docx4j ``198d015a3``, with the constants of ``ExtractionInspection`` it uses.
It checks an :class:`Engine`'s extraction of the suite's five fixtures
(``tests/fixtures/text-extraction/phase0/``, copied from docx4j at that
commit) against the suite's tables and goldens. ``GOLDENS.md`` and
``INSPECTION.md`` beside the suite are its contract. In order:

* each token row of ``expected-tokens.tsv`` in the three views, under its
  default scenario and its option column's, then under the supplemental
  scenarios that apply to its case;
* each row of ``expected-findings.tsv`` (warnings, metadata and inspection
  findings, an ``inspect`` row in ``accepted`` only) and of
  ``expected-order.tsv``;
* each fixture's whole inspection (REQ-1204), the HTML reader's fixture's,
  and REQ-903's counts under the six outputs;
* each golden of ``goldens.tsv``, compared after REQ-1201's normalisation.

A check **pends** when the engine does not advertise every statement it needs
(:meth:`Engine.implemented`): the row's statements and its scenario's, and for
a golden ``html reader absent``. The report is a TSV in docx4j's form
(:data:`HEADER`), one row per check.

**The one departure** (CR-011 section 11): docx4j's run stops at the first
check that fails, since every check docx4j makes passes, so its report has
``PASS`` and ``PENDING`` rows only. This port records a failing check as a
``FAIL`` row with its reason and goes on, and an engine's exception fails the
checks that read its output, so that today's readers can be measured.
``test_conformance.py`` holds the run to the committed report,
``tests/text/conformance.tsv``, which this module rewrites when run::

    .venv-fork/bin/python tests/text/conformance.py      # the adapter, the report rewritten

Every FAIL reason is this port's own (docx4j writes none): its assertion
message without the context the row's columns already give.
"""

from __future__ import annotations

import collections
import functools
import re
import sys
import unicodedata
from abc import ABC, abstractmethod
from collections.abc import Callable, Collection, Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import NamedTuple

__all__ = [
    "FIELDS",
    "HEADER",
    "HTML_READER_ABSENT",
    "HTML_READER_PRESENT",
    "KEY",
    "LIMITATIONS",
    "OUTPUTS",
    "PHASE0",
    "REPORT",
    "Check",
    "CheckFailed",
    "Engine",
    "Inspection",
    "Output",
    "Run",
    "compare_inspection",
    "count",
    "image_destinations",
    "main",
    "normalize",
    "phase",
    "read_report",
    "requirements",
    "row_requirements",
    "run",
    "scenario_requirements",
    "table",
    "text_requirements",
    "verify",
    "verify_golden",
    "verify_order",
]

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
#: The suite as copied from docx4j at ``198d015a3`` (``tests/README.md``).
SUITE = ROOT / "tests" / "fixtures" / "text-extraction"
#: The fixtures, the tables and the goldens.
PHASE0 = SUITE / "phase0"
#: The committed report: the adapter's run, which ``test_conformance.py`` holds.
REPORT = HERE / "conformance.tsv"

#: The class path's capabilities (CR-053, WD 0.27): whether an HTML reader is
#: present, a parser of the [HTML] algorithm and a reader of its tree as
#: blocks. Not options: an engine advertises the one it can check.
HTML_READER_PRESENT = "html reader present"
HTML_READER_ABSENT = "html reader absent"

# ExtractionInspection's constants -------------------------------------------

INSPECTION_TABLE = "expected-inspection.tsv"
INSPECTION_STATUS = "expected-inspection-status.tsv"
#: An inspector with an HTML reader, on the fixture whose chunks it reads.
HTML_TABLE = "expected-inspection-html.tsv"
HTML_STATUS = "expected-inspection-html-status.tsv"
#: A finding's identity: its kind and its stored location.
KEY = ("kind", "chain", "part", "path", "start", "end_path", "end")
#: The fields compared for a finding whose identity matched.
FIELDS = (
    "address",
    "visible",
    "reveal",
    "plain",
    "markdown",
    "blocks",
    "hidden",
    "embedded",
    "comments",
    "excerpt_length",
    "excerpt_cut",
)
#: What an inspection row needs: WD section 9, less REQ-903's counts, which
#: have their own check.
INSPECTOR_STATEMENTS = "REQ-901, REQ-902, REQ-904, REQ-905, REQ-906, REQ-907"
#: The six outputs the inspection table groups each finding under.
OUTPUTS = ("plain", "markdown", "blocks", "hidden", "embedded", "comments")
#: What every inspector says it does not detect (REQ-904), by id.
LIMITATIONS = (
    "layout.covered",
    "layout.position",
    "layout.clipped",
    "background.image",
    "font.careful_swap",
    "payload.opaque",
    "chart.workbook",
    "intent",
)

# the run ----------------------------------------------------------------------

#: The three views every token, findings and order row is read in.
VIEWS = ("accepted", "original", "markup")
#: The supplemental scenarios: combined capabilities, no-fetch fields and
#: character retention.
SUPPLEMENTAL = (
    "embedded: true",
    "altchunks: false",
    "fields: update",
    HTML_READER_PRESENT,
    "metadata: true",
    "field_codes: true",
)
#: The report's header, docx4j's.
HEADER = "fixture\tcase\titem\tview\tscenario\tstatus\treason"

#: A table row, by column name.
Row = Mapping[str, str]


class Output(NamedTuple):
    """An extraction: the text, and the findings by ``channel<TAB>kind<TAB>selector``.

    A finding's value is a count or a metadata value, as a string.
    """

    text: str
    findings: Mapping[str, str]


@dataclass(frozen=True)
class Inspection:
    """An inspection as the harness compares it with ``expected-inspection.tsv``.

    Attributes:
        findings: each finding's columns, named as the table names them: its
            :data:`KEY` and :data:`FIELDS` at least.
        status: ``COMPLETE`` or ``PARTIAL``.
        skipped: ``reason@chain|part|path`` entries.
        undecided: as `skipped`.
        limitations: REQ-904's statements, each starting with its id.
    """

    findings: Sequence[Mapping[str, str]]
    status: str
    skipped: Sequence[str] = ()
    undecided: Sequence[str] = ()
    limitations: Sequence[str] = ()


class Engine(ABC):
    """docx4j's ``ExtractionConformance.Engine``: what the harness runs.

    An engine advertises only the statements it implements, so that every row
    it cannot meet stays visible as pending. :meth:`inspect` and :meth:`counts`
    are an inspector's (CR-050); an engine that does not advertise REQ-901 to
    907 never has them called, as in docx4j, whose defaults throw.
    """

    @abstractmethod
    def implemented(self) -> Collection[str]:
        """The REQ ids, capabilities (``CR-047``) and class-path scenarios advertised."""

    @abstractmethod
    def extract(self, fixture: Path, view: str, scenario: str, format: str) -> Output:
        """One fixture read in one view, under one scenario, as ``txt`` or ``md``."""

    def inspect(self, fixture: Path) -> Inspection:
        """A fixture's inspection (WD section 9)."""
        raise NotImplementedError("not an inspector")

    def counts(self, fixture: Path, output: str) -> Mapping[str, str]:
        """REQ-903's counts of a fixture's result under one of :data:`OUTPUTS`."""
        raise NotImplementedError("not an inspector")


class CheckFailed(AssertionError):
    """A check that did not hold; the message is the report row's reason."""


# the tables -------------------------------------------------------------------

_LINE = re.compile(r"\r\n|\r|\n")


def table(path: Path) -> list[dict[str, str]]:
    """A TSV with a header row, as Java's ``Files.readAllLines`` splits it.

    Empty lines are skipped; a row with another number of columns than the
    header is an error of the suite, not of an engine.
    """
    lines = _LINE.split(path.read_bytes().decode("utf-8"))
    header = lines[0].split("\t")
    rows = []
    for number, line in enumerate(lines[1:], start=2):
        if not line:
            continue
        values = line.split("\t")
        if len(values) != len(header):
            raise ValueError(f"{path}:{number}: column count")
        rows.append(dict(zip(header, values, strict=True)))
    return rows


# REQ-1201 ---------------------------------------------------------------------

_MARKS = re.compile(r"\{\+\+|\+\+\}|\{--|--\}|\{>>|<<\}")
#: Java's ``\s`` (ASCII by default) and the characters REQ-1201 removes.
_REMOVED = frozenset(" \t\n\x0b\f\r*_\\{}$")


def _removed(character: str) -> bool:
    return character in _REMOVED or unicodedata.category(character) in ("Zs", "Zl", "Zp")


@functools.lru_cache(maxsize=4096)
def normalize(value: str) -> str:
    """REQ-1201's normalisation: CriticMarkup's brackets, then ``[\\s\\p{Z}*_\\\\{}$]``, removed.

    Zero-width spaces, tag characters, both hyphens and Unicode composition
    are kept.
    """
    return "".join(c for c in _MARKS.sub("", value) if not _removed(c))


def count(text: str, marker: str, suite: Iterable[str] = ()) -> int:
    """Occurrences of a marker, not counting those inside a longer marker token of the suite.

    After normalisation removes underscores and whitespace, ``FLD19_RESULT``
    is also found inside ``FLD19_RESULT_TWO``, and ``MCE8_OUTER`` inside
    ``MCE8_OUTER_FALLBACK``.
    """
    haystack, needle = normalize(text), normalize(marker)
    if not needle:
        raise ValueError("Empty marker")
    covered = [False] * len(haystack)
    for other in suite:
        longer = normalize(other)
        if len(longer) <= len(needle) or needle not in longer:
            continue
        at = haystack.find(longer)
        while at >= 0:
            covered[at : at + len(longer)] = [True] * len(longer)
            at = haystack.find(longer, at + len(longer))
    result = 0
    at = haystack.find(needle)
    while at >= 0:
        if not covered[at]:
            result += 1
        at = haystack.find(needle, at + len(needle))
    return result


_OPENING = {"ins": "{++", "del": "{--", "cmt": "{>>"}
_CLOSING = {"ins": "++}", "del": "--}", "cmt": "<<}"}


def _equal(expected: object, actual: object, context: str = "") -> None:
    if expected != actual:
        prefix = f"{context}: " if context else ""
        raise CheckFailed(f"{prefix}expected {expected}, got {actual}")


def verify(text: str, token: str, expected: str, suite: Iterable[str] = ()) -> None:
    """A token row's expectation: an exact count, or CriticMarkup kinds (``del+ins``).

    A kind is its marking's total count and one occurrence inside a marking of
    that kind.
    """
    suite = tuple(suite)
    if re.fullmatch(r"[0-9]+", expected):
        _equal(int(expected), count(text, token, suite))
        return
    kinds = expected.split("+")
    _equal(len(kinds), count(text, token, suite), "total")
    for kind in kinds:
        if kind not in _OPENING:
            raise ValueError(f"Unknown marking: {kind}")
        pattern = re.escape(_OPENING[kind]) + "(.*?)" + re.escape(_CLOSING[kind])
        marked = sum(count(m.group(1), token, suite) for m in re.finditer(pattern, text, re.DOTALL))
        _equal(1, marked, f"marking {kind}")


def verify_order(text: str, markers: str) -> None:
    """An order row: each marker once, each after the one before it."""
    normalized = normalize(text)
    after = 0
    for token in re.split(r"\s*>\s*", markers):
        marker = normalize(token)
        _equal(1, count(text, token), f"order marker {token}")
        position = normalized.find(marker)
        if position < after:
            raise CheckFailed(f"wrong order at {token}")
        after = position + len(marker)


_IMAGE = re.compile(r"!\[((?:\\.|[^\]\\])*)\]\([^)]*\)")


def image_destinations(markdown: str) -> str:
    """Markdown with each image's destination written ``…``.

    REQ-703 writes ``![alt](…)`` and leaves the destination to the
    implementation (docx4j writes a data URI, or a file's name), so a
    golden's is not compared.
    """
    return _IMAGE.sub(lambda m: "![" + m.group(1) + "](\u2026)", markdown)


def verify_golden(golden: Path, actual: str) -> None:
    """A golden against the text, both normalised (REQ-1201).

    The reason counts the golden's lines the text lacks and the lines it has
    that the golden does not, as multisets: this port's, docx4j's being the
    whole of both normalised strings.
    """
    expected = golden.read_bytes().decode("utf-8")
    if golden.name.endswith(".md"):
        expected, actual = image_destinations(expected), image_destinations(actual)
    if normalize(expected) == normalize(actual):
        return
    want = collections.Counter(expected.rstrip("\n").split("\n"))
    got = collections.Counter(actual.rstrip("\n").split("\n"))
    missing, extra = sum((want - got).values()), sum((got - want).values())
    raise CheckFailed(f"normalised text differs: {missing} lines of the golden missing, {extra} extra")


# the statements a check needs -------------------------------------------------


def requirements(value: str) -> list[str]:
    """A statements column: ``REQ-…`` ids, ``, `` apart, in order, each once."""
    return list(dict.fromkeys(re.split(r",\s*", value)))


#: The statement each WD section 10 option belongs to.
_OPTION_STATEMENTS = {
    "text_boxes": "REQ-518",
    "notes": "REQ-503",
    "note_references": "REQ-503",
    "comments": "REQ-504",
    "hidden": "REQ-507",
    "field_codes": "REQ-505",
    "fields": "REQ-506",
    "placeholders": "REQ-508",
    "toc": "REQ-520",
    "bindings": "REQ-509",
    "list_labels": "REQ-510",
    "charts": "REQ-511",
    "smartart": "REQ-511",
    "embedded": "REQ-512",
    "altchunks": "REQ-513",
    "alt_text": "REQ-514",
    "metadata": "REQ-515",
    "symbols": "REQ-601",
    "math": "REQ-604",
    "tables": "REQ-605",
    "markup_detail": "REQ-702",
    "page_markers": "REQ-708",
    "layout": "REQ-709",
}


def scenario_requirements(scenario: str | None) -> list[str]:
    """The statements a scenario's options belong to (WD section 10), which the engine needs too.

    Options may be combined, ``; `` apart (``embedded: true; hidden: true``):
    each one's statement. The two HTML reader scenarios need REQ-513 and the
    capability itself.
    """
    if scenario is None or scenario == "default":
        return []
    if scenario in (HTML_READER_PRESENT, HTML_READER_ABSENT):
        return ["REQ-513", scenario]
    if scenario == "markdown":
        return ["REQ-703"]
    if ";" in scenario:
        options = (scenario_requirements(option.strip()) for option in scenario.split(";"))
        return list(dict.fromkeys(r for statements in options for r in statements))
    name, _, value = (part.strip() for part in scenario.partition(":"))
    if name == "include":
        if "glossary" in value:
            return ["REQ-516"]
        return ["REQ-504"] if "comments" in value else ["REQ-502"]
    if name not in _OPTION_STATEMENTS:
        raise ValueError(f"Unknown scenario: {scenario}")
    return [_OPTION_STATEMENTS[name]]


def row_requirements(row: Row) -> list[str]:
    """A findings, order or golden row's statements and its scenario's, as a token row has."""
    return list(dict.fromkeys(requirements(row["statements"]) + scenario_requirements(row["scenario"])))


def text_requirements(row: Row, format: str) -> list[str]:
    """A token row's statements, as the text checks need them.

    REQ-902 is the inspector's, which an extractor does not claim; case 34's
    axis title needs ``CR-047``; case 21 in plain text needs no REQ-703; case
    25's text boxes are phase 1, only its legacy WordArt needing REQ-511.
    """
    statements = requirements(row["statements"])
    drop = {"REQ-902"}
    if row["case"] == "21" and format == "txt":
        drop.add("REQ-703")
    if row["case"] == "25" and not row["token"].startswith("WA25"):
        drop.add("REQ-511")
    statements = [s for s in statements if s not in drop]
    if row["token"] == "CHT34_AXIS_TITLE" and "CR-047" not in statements:
        statements.append("CR-047")
    return statements


_PHASE_2 = frozenset({506, 509, 511, 512, 513, 514, 515, 516, 604, 605, 708, 709})


def phase(statements: Iterable[str]) -> int:
    """The docx4j CR-046 phase a set of statements belongs to; 9 for the inspector (CR-050)."""
    result = 1
    for statement in statements:
        # a dependency (CR-047), a class path, or a ruling still to come
        if not statement.startswith("REQ-"):
            result = max(result, 2)
            continue
        number = int(statement[4:])
        if 901 <= number <= 907:
            return 9
        if number in _PHASE_2:
            result = max(result, 2)
        if number in (703, 704, 705):
            result = max(result, 3)
    return result


# the inspection ---------------------------------------------------------------


def _key(finding: Mapping[str, str]) -> str:
    return " | ".join(finding.get(k, "") for k in KEY)


def _occurrences(findings: Iterable[Mapping[str, str]], names: Sequence[str]) -> dict[str, list[str]]:
    """Location key to the sorted field tuples of the findings there: a multiset."""
    out: dict[str, list[str]] = collections.defaultdict(list)
    for finding in findings:
        out[_key(finding)].append(" | ".join(finding.get(k, "") for k in names))
    return {key: sorted(values) for key, values in sorted(out.items())}


def compare_inspection(
    expected: Sequence[Mapping[str, str]],
    status: Mapping[str, str],
    actual: Inspection,
    fields: Sequence[str] = FIELDS,
) -> None:
    """An inspection against its table, by occurrence (REQ-1204).

    The findings as multisets of (kind, location), so that a finding missing,
    one unexpected and one of two at a location all fail; then each matched
    finding's fields, the status, what was skipped and undecided, and the
    limitations every inspector states.
    """
    want, got = _occurrences(expected, fields), _occurrences(actual.findings, fields)
    problems = []
    for key in sorted(set(want) | set(got)):
        w, g = want.get(key, []), got.get(key, [])
        if len(w) != len(g):
            problems.append(f"{'missing' if len(w) > len(g) else 'unexpected'} {abs(len(w) - len(g))}: {key}")
        elif w != g:
            problems.append(f"fields of {key}")
    if problems:
        raise CheckFailed(f"inspection differs ({len(problems)}): {problems[0]}")
    _equal(status["status"], actual.status, "status")
    _equal(status["skipped"], "; ".join(sorted(actual.skipped)), "skipped")
    _equal(status["undecided"], "; ".join(sorted(actual.undecided)), "undecided")
    for limitation in LIMITATIONS:
        if not any(stated == limitation or stated.startswith(limitation + ":") for stated in actual.limitations):
            raise CheckFailed(f"limitation not stated: {limitation}")


# the report -------------------------------------------------------------------


@dataclass(frozen=True)
class Check:
    """One row of the report: a check and its outcome (``PASS``, ``FAIL`` or ``PENDING``)."""

    fixture: str
    case: str
    item: str
    view: str
    scenario: str
    status: str
    reason: str = ""

    @property
    def key(self) -> tuple[str, str, str, str, str]:
        """What identifies a check across runs: everything but its outcome."""
        return (self.fixture, self.case, self.item, self.view, self.scenario)

    def line(self) -> str:
        """The row as the report writes it."""
        return "\t".join((*self.key, self.status, self.reason))


def _reason(text: str) -> str:
    """A reason on one line, with no tab: the report is a TSV."""
    return " ".join(text.split())


@dataclass
class Run:
    """One run of the harness over the suite: the checks in order, and the outputs read."""

    engine: Engine
    directory: Path = PHASE0
    tokens: tuple[str, ...] = ()
    checks: list[Check] = field(default_factory=list)
    _outputs: dict[tuple[str, str, str, str], Output | str] = field(default_factory=dict)

    def _add(self, row: Row, item: str, view: str, scenario: str, status: str, reason: str = "") -> None:
        self.checks.append(Check(row["fixture"], row["case"], item, view, scenario, status, _reason(reason)))

    def ready(self, row: Row, item: str, view: str, scenario: str, statements: Iterable[str]) -> bool:
        """Whether the engine advertises every statement a check needs; if not, the check pends."""
        statements = list(dict.fromkeys(statements))
        implemented = set(self.engine.implemented())
        missing = [s for s in statements if s not in implemented]
        if not missing:
            return True
        n = phase(statements)
        owner = "the inspector (CR-050)" if n == 9 else f"phase {n}"
        self._add(row, item, view, scenario, "PENDING", f"{owner}; {', '.join(missing)}")
        return False

    def pend(self, row: Row, item: str, view: str, scenario: str, reason: str) -> None:
        """A check left pending for a reason of its own (a rule not yet in the WD)."""
        self._add(row, item, view, scenario, "PENDING", reason)

    def passed(self, row: Row, item: str, view: str, scenario: str) -> None:
        """A check that held."""
        self._add(row, item, view, scenario, "PASS")

    def failed(self, row: Row, item: str, view: str, scenario: str, reason: str) -> None:
        """A check that did not hold (this port's; docx4j stops)."""
        self._add(row, item, view, scenario, "FAIL", reason)

    def output(self, row: Row, view: str, scenario: str, format: str) -> Output:
        """The engine's extraction, read once per (fixture, view, scenario, format).

        Raises:
            CheckFailed: the engine raised; every check over this output fails
                with its exception.
        """
        key = (row["fixture"], view, scenario, format)
        if key not in self._outputs:
            try:
                self._outputs[key] = self.engine.extract(self.directory / row["fixture"], view, scenario, format)
            except Exception as error:  # noqa: BLE001 - an engine's failure is a check's, as in docx4j
                self._outputs[key] = f"{type(error).__name__}: {error}"
        out = self._outputs[key]
        if isinstance(out, str):
            raise CheckFailed(out)
        return out

    def outcome(self, row: Row, item: str, view: str, scenario: str, check: Callable[[], None]) -> None:
        """Run one ready check and record what it gave.

        Only :class:`CheckFailed` is a failure: anything else is an error of
        the suite or of this harness, and stops the run.
        """
        try:
            check()
        except CheckFailed as failure:
            self.failed(row, item, view, scenario, str(failure))
        else:
            self.passed(row, item, view, scenario)

    def inspection(self, fixture: str) -> Inspection:
        """The engine's inspection of a fixture; its exception is the check's failure."""
        try:
            return self.engine.inspect(self.directory / fixture)
        except Exception as error:  # an engine's failure is a check's, as in docx4j
            raise CheckFailed(f"{type(error).__name__}: {error}") from error

    def counts(self, fixture: str, output: str) -> dict[str, str]:
        """The engine's REQ-903 counts; its exception is the check's failure."""
        try:
            return dict(self.engine.counts(self.directory / fixture, output))
        except Exception as error:  # an engine's failure is a check's, as in docx4j
            raise CheckFailed(f"{type(error).__name__}: {error}") from error

    # -- the summary ---------------------------------------------------------

    def count_of(self, status: str) -> int:
        """How many checks have that outcome."""
        return sum(1 for c in self.checks if c.status == status)

    @property
    def goldens(self) -> tuple[int, int]:
        """(passing, all) of the golden checks."""
        goldens = [c for c in self.checks if c.item.startswith("golden:")]
        return sum(1 for c in goldens if c.status == "PASS"), len(goldens)

    def summary(self) -> str:
        """docx4j's line, with the failures and the goldens this port adds."""
        passing, goldens = self.goldens
        return (
            f"CR-046: {self.count_of('PASS')} assertions passed, {self.count_of('FAIL')} failed, "
            f"{self.count_of('PENDING')} pending of {len(self.checks)}; goldens {passing} of {goldens}"
        )

    def report(self) -> str:
        """The report, as docx4j's run writes it, with this port's ``FAIL`` rows."""
        return "\n".join([HEADER, *(c.line() for c in self.checks)]) + "\n"


def read_report(path: Path = REPORT) -> list[Check]:
    """A report written by :meth:`Run.report`."""
    return [Check(**row) for row in table(path)]


# the checks -------------------------------------------------------------------


def _check_token(run: Run, row: Row, view: str, scenario: str, format: str, expected: str) -> None:
    statements = text_requirements(row, format) + scenario_requirements(scenario)
    if not run.ready(row, row["token"], view, scenario, statements):
        return

    def check() -> None:
        text = run.output(row, view, scenario, format).text
        verify(text, row["token"], expected, run.tokens)
        if row["token"] == "TAG48_MARK":
            _equal(True, "TAG48_MARK\U000e0041\U000e0042" in text, "REQ-602 tag characters retained")

    run.outcome(row, row["token"], view, scenario, check)


def _applies(scenario: str, row: Row) -> bool:
    """Whether a supplemental scenario reads a token row's case."""
    n = int(row["case"])
    return (
        scenario == "embedded: true" and (39 <= n <= 42 or n in (53, 57))
        or scenario == "altchunks: false" and (43 <= n <= 47 or n == 36 or row["token"] == "CHK63_SCRIPTED")
        or scenario == "fields: update" and 49 <= n <= 51
        or scenario == HTML_READER_PRESENT and n == 47
        or scenario in ("metadata: true", "field_codes: true", "embedded: true") and 50 <= n <= 51
    )


def _tokens(run: Run, rows: Sequence[Row]) -> None:
    for row in rows:
        for view in VIEWS:
            # a row gated on an HTML reader (an HTML or MHT chunk) runs its default scenario without one
            scenario = HTML_READER_ABSENT if row["option"] == HTML_READER_PRESENT else "default"
            _check_token(run, row, view, scenario, "txt", row[view])
            option = row["option"]
            if option and (option != "comments: true" or view == "markup"):
                _check_token(run, row, view, option, "md" if option == "markdown" else "txt", row["with_option"])
    for scenario in SUPPLEMENTAL:
        for row in rows:
            if not _applies(scenario, row) or row["option"] == scenario:
                continue
            for view in VIEWS:
                expected = "0" if scenario == "altchunks: false" else row[view]
                _check_token(run, row, view, scenario, "txt", expected)


def _findings(run: Run) -> None:
    for row in table(run.directory / "expected-findings.tsv"):
        item = f"{row['channel']}:{row['kind']}:{row['selector']}"
        # an inspection is the package's, whatever the view (CR-050): its rows are checked once
        for view in ("accepted",) if row["channel"] == "inspect" else VIEWS:
            if not run.ready(row, item, view, row["scenario"], row_requirements(row)):
                continue

            def check(row: Row = row, view: str = view, item: str = item) -> None:
                findings = run.output(row, view, row["scenario"], "txt").findings
                _equal(row["expected"], findings.get(f"{row['channel']}\t{row['kind']}\t{row['selector']}"))

            run.outcome(row, item, view, row["scenario"], check)


def _orders(run: Run) -> None:
    path = run.directory / "expected-order.tsv"
    if not path.exists():
        return
    for row in table(path):
        for view in VIEWS:
            if not run.ready(row, "order", view, row["scenario"], row_requirements(row)):
                continue

            def check(row: Row = row, view: str = view) -> None:
                verify_order(run.output(row, view, row["scenario"], "txt").text, row["tokens"])

            run.outcome(row, "order", view, row["scenario"], check)


def _by_fixture(rows: Iterable[Row]) -> dict[str, list[Row]]:
    out: dict[str, list[Row]] = collections.defaultdict(list)
    for row in rows:
        out[row["fixture"]].append(row)
    return out


def _inspections(run: Run) -> None:
    """Each fixture's inspection by occurrence, the HTML reader's, and REQ-903's counts.

    The table is an inspector's with no HTML reader (CR-053 phase 3), so the
    comparisons and the counts need ``html reader absent``.
    """
    path = run.directory / INSPECTION_TABLE
    if not path.exists():
        return
    expected = _by_fixture(table(path))
    statuses = table(run.directory / INSPECTION_STATUS)
    for status in statuses:
        fixture = status["fixture"]
        statements = INSPECTOR_STATEMENTS + ", REQ-1204"
        row = {"fixture": fixture, "case": "0", "statements": statements}
        if not run.ready(row, "inspection", "-", "default", [*requirements(statements), HTML_READER_ABSENT]):
            continue

        def check(fixture: str = fixture, status: Row = status) -> None:
            compare_inspection(expected.get(fixture, []), status, run.inspection(fixture))

        run.outcome(row, "inspection", "-", "default", check)
    _html_inspections(run)
    for status in statuses:
        fixture = status["fixture"]
        row = {"fixture": fixture, "case": "0", "statements": "REQ-903"}
        for output in OUTPUTS:
            if not run.ready(row, "counts", "-", output, ["REQ-903", HTML_READER_ABSENT]):
                continue

            def check(fixture: str = fixture, status: Row = status, output: str = output) -> None:
                want = {
                    group: str(sum(1 for f in expected.get(fixture, []) if f[output] == group))
                    for group in ("unseen", "omitted", "active", "info")
                }
                want["status"] = status["status"]
                _equal(want, run.counts(fixture, output), f"counts under {output}")

            run.outcome(row, "counts", "-", output, check)


def _html_inspections(run: Run) -> None:
    """An inspector with an HTML reader, on the fixture whose chunks it reads (CR-053 phase 3).

    ``parsed`` is compared with the other fields; a row resting on a rule not
    yet ruled on (its ``pending``) is left out on both sides and pends.
    """
    path = run.directory / HTML_TABLE
    if not path.exists():
        return
    expected = _by_fixture(table(path))
    for status in table(run.directory / HTML_STATUS):
        fixture = status["fixture"]
        statements = INSPECTOR_STATEMENTS + ", REQ-1204, REQ-513"
        row = {"fixture": fixture, "case": "0", "statements": statements}
        needs = [*requirements(statements), HTML_READER_PRESENT]
        if not run.ready(row, "inspection", "-", HTML_READER_PRESENT, needs):
            continue
        checked, pending = [], set()
        for finding in expected.get(fixture, []):
            if not finding["pending"]:
                checked.append(finding)
            else:
                pending.add(_key(finding))
                item = f"inspect:{finding['kind']}:{finding['selector']}"
                run.pend(finding, item, "-", HTML_READER_PRESENT, finding["pending"])

        def check(
            fixture: str = fixture, status: Row = status, checked: list[Row] = checked, pending: set[str] = pending
        ) -> None:
            actual = run.inspection(fixture)
            found = [f for f in actual.findings if _key(f) not in pending]
            others = Inspection(found, actual.status, actual.skipped, actual.undecided, actual.limitations)
            compare_inspection(checked, status, others, (*FIELDS, "parsed"))

        run.outcome(row, "inspection", "-", HTML_READER_PRESENT, check)


def _goldens(run: Run) -> None:
    for row in table(run.directory / "goldens.tsv"):
        item = "golden:" + row["golden"]
        # a golden is cut without an HTML reader (GOLDENS.md), so with one it pends
        statements = row_requirements(row)
        if row["scenario"] == "default":
            statements.append(HTML_READER_ABSENT)
        if not run.ready(row, item, row["view"], row["scenario"], statements):
            continue

        def check(row: Row = row) -> None:
            text = run.output(row, row["view"], row["scenario"], row["format"]).text
            verify_golden(run.directory / row["golden"], text)

        run.outcome(row, item, row["view"], row["scenario"], check)


def run(engine: Engine, directory: Path = PHASE0) -> Run:
    """The suite over an engine: every check, in docx4j's order, with its outcome."""
    tokens = table(directory / "expected-tokens.tsv")
    result = Run(engine, directory, tuple(dict.fromkeys(row["token"] for row in tokens)))
    _tokens(result, tokens)
    _findings(result)
    _orders(result)
    _inspections(result)
    _goldens(result)
    return result


def main(argv: Sequence[str] | None = None) -> int:
    """Run the adapter over the suite, print the summary and write the report.

    ``conformance.py [report.tsv]``: the committed report by default.
    """
    arguments = list(sys.argv[1:] if argv is None else argv)
    report = Path(arguments[0]) if arguments else REPORT
    from adapter import TextsAdapter

    result = run(TextsAdapter())
    report.write_text(result.report(), encoding="utf-8", newline="\n")
    print(f"{result.summary()}; {report}")
    return 0


if __name__ == "__main__":
    # the package and the adapter, as tests/text/conftest.py puts them on the path
    sys.path[:0] = [str(ROOT), str(HERE)]
    raise SystemExit(main())
