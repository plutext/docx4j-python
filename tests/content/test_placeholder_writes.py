"""Writing into a content control showing its placeholder: CR-019 Phase A, the port of core-ts's CR-002 section 63.

    .venv-fork/bin/python -m pytest tests/content/test_placeholder_writes.py

docx4j-core-ts ran Word 365 checks 63 and 64 on what Word does with every text
write into a content control showing its placeholder (its CR-002 sections 63.2
and 63.5; Word 365 16.0.20326.20232, 2026-10-24), and implemented the answers at
``020c039``. Its ``test/placeholders.test.mjs`` replays them through its Office
JS shim; this package has none, so ``tests/office_js/checks.py`` ports the two
Script Lab snippets step for step (:func:`run_check_63`, :func:`run_check_64`),
and this file holds this package's run to Word's, case by case, through
core-ts's ``digest`` and ``view``, then the markup of each region against Word's
save after the same run (rsids, paraIds, ``w:sdtEndPr`` and revision ``w:id``
values left out, as CR-016's check 58 comparison leaves them).

* **Check 63** is the bare form (no ``w:placeholder``, no glossary): what
  docx4j's binding and this package write.
* **Check 64** is Word's form (``w:placeholder`` naming Word's default entry,
  check 34's glossary), with controls Word makes and bound controls' nodes.

What still differs is pinned, with its reason: the bare form's look (check 63's
``PlaceholderText``, ``w:color auto`` and ``w:rPrChange``, contradicted by
core-ts's checks 58 and 67: CR-019 Q1), the controls ``insert_content_control``
makes (Phase B's, D8), and in Word's saves D6 (b)'s level and Word's run
boundaries. Then core-ts's tests that are no replay: D3's markup, accepted and
rejected; D6's level; D7's nodes in both forms. And this package's own: every
route, the control's own ranges, Q4, the nodes.
"""

from __future__ import annotations

import hashlib
import importlib.util
import io
import json
import re
import zipfile
from typing import Any

import pytest
from conftest import ROOT
from lxml import etree

from docx4j_py import create_package, load
from docx4j_py.wml import to_xml

OFFICE_JS = ROOT / "tests" / "office_js"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_spec = importlib.util.spec_from_file_location("office_js_checks", OFFICE_JS / "checks.py")
checks = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(checks)

#: Word's default placeholder text and entry (check 64's controls name it).
PLACEHOLDER = "Click or tap here to enter text."
DEFAULT_ENTRY = "DefaultPlaceholder_-1854013440"

#: The files, as core-ts 020c039 has them: each check's file (git-ignored there, written by its build.mjs on
#: 2026-10-24 before Word's run, and read here as Word's save of it as opened reads), Word's two JSONs and Word's
#: two saves (its manifest.json's SHA-1s).
SHA1 = {
    "63-placeholders.docx": "788396d96be54ac1b4f825d8c1de5d1eb9fb93c3",
    "63-placeholders-word365-report.json": "a84e8987b3d79653ea91136ef1a120e61e401418",
    "63-placeholders-word365-run.json": "8f0404a4b301669934d13315e295e3dc44e0e284",
    "63-placeholders-word365.docx": "b8dc3e3b7b272a7c7bf0dbf7fdb76336ac78a3e8",
    "63-run-word365.docx": "6c118dcce476b37e7d3f61ed252af0eec6c325b2",
    "64-placeholders-word.docx": "fc2e15bfb4c50e1111cfb834cb132e8f36fc8450",
    "64-placeholders-word-word365-report.json": "eb265f08a79c29f71d658e561af7d60eca378fbb",
    "64-placeholders-word-word365-run.json": "4da651a179bfbf0cc06df6c9d20861f30bff7269",
    "64-opened-word365.docx": "bff704e268525a444db9732af7cd44df6d1404a0",
    "64-run-word365.docx": "62220a054f033c842ed4ba1f7711c9e0f314a361",
}


def _cases_63() -> list[dict[str, Any]]:
    """Check 63's cases as ``build.mjs`` at core-ts ``020c039`` builds them (``check63Cases``)."""
    block = [
        "control-text-replace",
        "control-text-end",
        "control-paragraph-end",
        "content-text-replace",
        "content-text-end",
        "content-delete",
        "paragraph-text-replace",
        "paragraph-text-start",
        "paragraph-text-end",
        "paragraph-after",
        "paragraph-before",
        "paragraph-content-delete",
    ]
    run = [
        "control-text-replace",
        "control-text-end",
        "content-text-replace",
        "content-text-end",
        "content-delete",
        "whole-text-replace",
        "paragraph-text-replace",
        "paragraph-content-delete",
    ]
    bound = ["control-text-replace", "content-text-replace", "paragraph-text-replace"]
    out: list[dict[str, Any]] = []

    def add(op: str, shape: str, mode: str | None = None) -> None:
        out.append({"n": len(out) + 1, "op": op, "shape": shape, **({"mode": mode} if mode else {})})

    for mode in ("untracked", "tracked"):
        for ops, shape in ((block, "block"), (run, "run"), (bound, "block-bound"), (bound, "run-bound")):
            for op in ops:
                add(op, shape, mode)
    add("open", "run-empty")
    add("open", "block-empty")
    return out


def _cases_64() -> list[dict[str, Any]]:
    """Check 64's cases as ``build.mjs`` at core-ts ``020c039`` builds them (``check64Cases``)."""
    text = [
        "control-text-replace",
        "control-text-start",
        "control-text-end",
        "content-text-replace",
        "content-text-end",
    ]
    block = [*text, "control-paragraph-end", "paragraph-text-replace", "paragraph-text-start", "paragraph-text-end"]
    made_block = ["control-text-replace", "control-text-end", "content-text-replace", "paragraph-text-end"]
    made_run = ["control-text-replace", "control-text-end", "content-text-replace"]
    node_block = ["control-text-replace", "content-text-replace", "paragraph-text-replace"]
    node_run = ["control-text-replace", "content-text-replace"]
    out: list[dict[str, Any]] = []

    def add(op: str, shape: str, mode: str) -> None:
        out.append({"n": len(out) + 1, "op": op, "shape": shape, "mode": mode})

    for mode in ("untracked", "tracked"):
        for ops, shape in ((block, "block"), (text, "run"), (made_block, "made-block"), (made_run, "made-run")):
            for op in ops:
                add(op, shape, mode)
    for shape in ("block-bound", "block-value"):
        for op in node_block:
            add(op, shape, "untracked")
    for shape in ("run-bound", "run-value"):
        for op in node_run:
            add(op, shape, "untracked")
    return out


CASES_63 = _cases_63()
CASES_64 = _cases_64()


def word(name: str) -> dict:
    return json.loads((OFFICE_JS / name).read_text(encoding="utf-8"))


def digest(case: dict[str, Any]) -> str:
    """A region after the case as one line: ``<n> <control><flag><style> <where> <paragraphs>`` (core-ts's ``digest``).

    The control is ``C`` where the region's OOXML still holds a ``w:sdt`` (else
    ``x``); ``P`` where its ``w:sdtPr`` has ``w:showingPlcHdr``; ``S`` where a run
    in its content has the Placeholder Text style among its own properties
    (``s`` where only deleted runs have it). Where ``NEW<n>`` landed: ``in``,
    ``out`` (outside the control), or ``-``. Then the region's paragraphs' text,
    ``|`` between, and the write's error code if it threw.
    """
    after = case.get("after") if isinstance(case.get("after"), dict) else {}
    xml = after.get("ooxml") if isinstance(after.get("ooxml"), str) else ""
    sdt = re.search(r"<w:sdt>([\s\S]*?)</w:sdt>", xml)
    pr = content = ""
    if sdt:
        found = re.search(r"<w:sdtPr>[\s\S]*?</w:sdtPr>", sdt.group(1))
        pr = found.group(0) if found else ""
        found = re.search(r"<w:sdtContent>([\s\S]*?)</w:sdtContent>", sdt.group(1))
        content = found.group(1) if found else ""
    # a run's own properties, not those a w:rPrChange records
    own = re.sub(r"<w:rPrChange\b[\s\S]*?</w:rPrChange>", "", content)
    runs = re.finditer(r"<w:r>(?:<w:rPr>([\s\S]*?)</w:rPr>)?<w:(t|delText)\b", own)
    styled = [m.group(2) for m in runs if '<w:rStyle w:val="PlaceholderText"/>' in (m.group(1) or "")]
    style = "S" if "t" in styled else "s" if styled else "-"
    token = f"NEW{case['n']}"
    where = "in" if token in content else "out" if token in xml else "-"
    error = ""
    if case.get("error"):
        found = re.match(r"^error: (\w+)", case["error"])
        error = " " + (found.group(1) if found else case["error"])
    paragraphs = after.get("paragraphs")
    paragraphs = "|".join(paragraphs) if isinstance(paragraphs, list) else str(paragraphs)
    flag = "P" if "showingPlcHdr" in pr else "-"
    return f"{case['n']} {'C' if sdt else 'x'}{flag}{style} {where} {paragraphs}{error}"


def view(case: dict[str, Any]) -> dict[str, Any]:
    """What both read (core-ts's ``view``): the digest, the control's type, text and paragraphs, what the write returned.

    ``placeholderText`` and ``showingPlaceholderText`` are left out, as core-ts
    leaves them: Word gives null or the glossary's text and no
    ``showingPlaceholderText``, and the reading is Phase B's (CR-019 Q2).
    """
    after = case.get("after") if isinstance(case.get("after"), dict) else {}
    control = after.get("control")
    if isinstance(control, dict):
        control = {key: control.get(key) for key in ("type", "text", "paragraphs")}
    return {"digest": digest(case), "control": control, "returned": (case.get("result") or {}).get("returned")}


#: Word 365's run of check 63, case by case, as `digest` reads it: core-ts's `WORD_RUN` (its CR-002 section 63.2).
WORD_RUN_63 = [
    "1 C-S in NEW1",
    "2 C-S in NEW2",
    "3 CPS in ",
    "4 C-S in NEW4",
    "5 C-S in NEW5",
    "6 CP- - ",
    "7 C-S in NEW7",
    "8 C-S in NEW8",
    "9 C-S in NEW9",
    "10 CPS out |NEW10",
    "11 CPS out NEW11|",
    "12 x-- - ",
    "13 C-S in PLAIN13A NEW13 PLAIN13B",
    "14 C-S in PLAIN14A NEW14 PLAIN14B",
    "15 C-S in PLAIN15A NEW15 PLAIN15B",
    "16 C-S in PLAIN16A NEW16 PLAIN16B",
    "17 CP- - PLAIN17A  PLAIN17B",
    "18 x-- out PLAIN18A NEW18 PLAIN18B",
    "19 x-- out NEW19",
    "20 x-- - ",
    "21 C-- in NEW21",
    "22 C-- in NEW22",
    "23 C-- in NEW23",
    "24 C-- in PLAIN24A NEW24 PLAIN24B",
    "25 C-- in PLAIN25A NEW25 PLAIN25B",
    "26 x-- out NEW26",
    "27 C-S in NEW27",
    "28 C-S in NEW28",
    "29 CPS in ",
    "30 C-S in NEW30",
    "31 C-S in NEW31",
    "32 CPs - ",
    "33 C-S in NEW33",
    "34 C-S in NEW34",
    "35 C-S in NEW35",
    "36 CPS out |NEW36",
    "37 CPS out NEW37|",
    "38 CPs - ",
    "39 C-S in PLAIN39A NEW39 PLAIN39B",
    "40 C-S in PLAIN40A NEW40 PLAIN40B",
    "41 C-S in PLAIN41A NEW41 PLAIN41B",
    "42 C-S in PLAIN42A NEW42 PLAIN42B",
    "43 CPs - PLAIN43A  PLAIN43B",
    "44 CPs out PLAIN44A NEW44 PLAIN44B",
    "45 CPs out NEW45",
    "46 CPs - ",
    "47 C-- in NEW47",
    "48 C-- in NEW48",
    "49 C-- in NEW49",
    "50 C-- in PLAIN50A NEW50 PLAIN50B",
    "51 C-- in PLAIN51A NEW51 PLAIN51B",
    "52 CPs out NEW52",
    "53 CP- - PLAIN53A  PLAIN53B",
    "54 CP- - ",
]


#: Word 365's run of check 64, as `digest` reads it: core-ts's `WORD_RUN_64` (its CR-002 section 63.5).
WORD_RUN_64 = [
    "1 C-- in NEW1",
    "2 C-- in NEW2",
    "3 C-- in NEW3",
    "4 C-- in NEW4",
    "5 C-- in NEW5",
    "6 C-- in |NEW6",
    "7 C-- in NEW7",
    "8 C-- in NEW8",
    "9 C-- in NEW9",
    "10 C-- in PLAIN10A NEW10 PLAIN10B",
    "11 C-- in PLAIN11A NEW11 PLAIN11B",
    "12 C-- in PLAIN12A NEW12 PLAIN12B",
    "13 C-- in PLAIN13A NEW13 PLAIN13B",
    "14 C-- in PLAIN14A NEW14 PLAIN14B",
    "15 C-- in NEW15",
    "16 C-- in NEW16",
    "17 C-- in NEW17",
    "18 C-- in NEW18",
    "19 C-- in PLAIN19ANEW19  PLAIN19B",
    "20 C-- in PLAIN20ANEW20  PLAIN20B",
    "21 C-- in PLAIN21ANEW21  PLAIN21B",
    "22 C-- in NEW22",
    "23 C-- in NEW23",
    "24 C-- in NEW24",
    "25 C-- in NEW25",
    "26 C-- in NEW26",
    "27 C-- in |NEW27",
    "28 C-- in NEW28",
    "29 C-- in NEW29",
    "30 C-- in NEW30",
    "31 C-- in PLAIN31A NEW31 PLAIN31B",
    "32 C-- in PLAIN32A NEW32 PLAIN32B",
    "33 C-- in PLAIN33A NEW33 PLAIN33B",
    "34 C-- in PLAIN34A NEW34 PLAIN34B",
    "35 C-- in PLAIN35A NEW35 PLAIN35B",
    "36 C-- in NEW36",
    "37 C-- in NEW37",
    "38 C-- in NEW38",
    "39 C-- in NEW39",
    "40 C-- in PLAIN40ANEW40  PLAIN40B",
    "41 C-- in PLAIN41ANEW41  PLAIN41B",
    "42 C-- in PLAIN42ANEW42  PLAIN42B",
    "43 C-- in NEW43",
    "44 C-- in NEW44",
    "45 C-- in NEW45",
    "46 C-- in NEW46",
    "47 C-- in NEW47",
    "48 C-- in NEW48",
    "49 C-- in PLAIN49A NEW49 PLAIN49B",
    "50 C-- in PLAIN50A NEW50 PLAIN50B",
    "51 C-- in PLAIN51A NEW51 PLAIN51B",
    "52 C-- in PLAIN52A NEW52 PLAIN52B",
]


#: Why a case differs from Word's.
Q1 = "the bare form's look: check 63's PlaceholderText and w:color auto, contradicted by checks 58 and 67 (CR-019 Q1)"
MADE = "a control insert_content_control made, which shows no placeholder here: Phase B's (CR-019 section 4.8, D8)"
ALONE = "a rule resting on check 63 alone, the commit after the first's (CR-019 section 10)"

#: Check 63's writes whose run differs from Word's: ``n: (why, this package's digest, its other reads or None)``.
#: Each is the bare form's look alone: the text written is a plain run here (``-`` where Word's ``S``), and tracked
#: the placeholder's deleted runs keep their style (``s``), where in a bound plain-text control Word took it off them
#: too, recorded in a ``w:rPrChange`` (cases 47 to 51; core-ts's D2 (a)).
DIFFERS_63: dict[int, tuple[str, str, dict[str, Any] | None]] = {
    **{n: (Q1, f"{n} C-- in NEW{n}", None) for n in (1, 2, 4, 5, 7, 8, 9)},
    **{n: (Q1, f"{n} C-- in PLAIN{n}A NEW{n} PLAIN{n}B", None) for n in (13, 14, 15, 16)},
    3: (Q1, "3 CP- in ", None),
    **{n: (Q1, f"{n} C-s in NEW{n}", None) for n in (27, 28, 30, 31, 33, 34, 35, 47, 48, 49)},
    **{n: (Q1, f"{n} C-s in PLAIN{n}A NEW{n} PLAIN{n}B", None) for n in (39, 40, 41, 42, 50, 51)},
    29: (Q1, "29 CPs in ", None),
}

#: Check 64's writes whose run differs from Word's: the controls made with insert_content_control, Phase B's.
DIFFERS_64: dict[int, tuple[str, str, dict[str, Any] | None]] = {
    19: (
        MADE,
        "19 C-- in NEW19 PLAIN19B",
        {"control": {"type": "RichText", "text": "NEW19", "paragraphs": ["NEW19 PLAIN19B"]}},
    ),
    20: (
        MADE,
        "20 C-- in PLAIN20A NEW20 PLAIN20B",
        {"control": {"type": "RichText", "text": "PLAIN20A NEW20", "paragraphs": ["PLAIN20A NEW20 PLAIN20B"]}},
    ),
    21: (
        MADE,
        "21 C-- in NEW21 PLAIN21B",
        {"control": {"type": "RichText", "text": "NEW21", "paragraphs": ["NEW21 PLAIN21B"]}},
    ),
    36: (MADE, "36 C-- in |NEW36", {"control": {"type": "RichText", "text": "NEW36", "paragraphs": ["", "NEW36"]}}),
    40: (
        MADE,
        "40 C-- in NEW40 PLAIN40B",
        {"control": {"type": "RichText", "text": "NEW40", "paragraphs": ["NEW40 PLAIN40B"]}},
    ),
    41: (
        MADE,
        "41 C-- in PLAIN41A NEW41 PLAIN41B",
        {"control": {"type": "RichText", "text": "PLAIN41A NEW41", "paragraphs": ["PLAIN41A NEW41 PLAIN41B"]}},
    ),
    42: (
        MADE,
        "42 C-- in NEW42 PLAIN42B",
        {"control": {"type": "RichText", "text": "NEW42", "paragraphs": ["NEW42 PLAIN42B"]}},
    ),
}


def run_63(mode: str = "run") -> tuple[Any, dict[str, Any]]:
    package = load(OFFICE_JS / "63-placeholders.docx")
    return package, checks.run_check_63(package, mode)


def run_64(mode: str = "run") -> tuple[Any, dict[str, Any]]:
    package = load(OFFICE_JS / "64-placeholders-word.docx")
    return package, checks.run_check_64(package, mode)


def _held(word_cases: list[dict], ours: list[dict], differs: dict[int, tuple[str, str, dict | None]]) -> list[int]:
    """The writes agreeing with Word's; each pinned one is asserted as it stands, its other reads Word's but those pinned."""
    agreeing = []
    for theirs, mine in zip(word_cases, ours, strict=True):
        n = theirs["n"]
        expected = view(theirs)
        if n in differs:
            _why, pinned, rest = differs[n]
            expected = {**expected, "digest": pinned, **(rest or {})}
        assert view(mine) == expected, n
        if n not in differs and "mode" in theirs:
            agreeing.append(n)
    return agreeing


# ---------------------------------------------------------------------------
# the files, Word's answers, and the port's agreement with the snippets
# ---------------------------------------------------------------------------


def test_the_files_are_core_tss_as_copied_and_word_365s_jsons_are_of_these_cases():
    manifest = dict(
        reversed(line.split("  ", 1))
        for line in (OFFICE_JS / "SHA1SUMS").read_text(encoding="utf-8").splitlines()
        if line
    )
    for name, sha1 in SHA1.items():
        assert manifest[name] == sha1, name
        assert hashlib.sha1((OFFICE_JS / name).read_bytes()).hexdigest() == sha1, name
    for stem, number, count, cases in (
        ("63-placeholders-word365", 63, 54, CASES_63),
        ("64-placeholders-word-word365", 64, 52, CASES_64),
    ):
        for mode in ("report", "run"):
            measured = word(f"{stem}-{mode}.json")
            assert [measured["check"], measured["mode"], measured["host"]["version"], measured["caseCount"]] == [
                number,
                mode,
                "16.0.20326.20232",
                count,
            ]
            assert [{k: c[k] for k in ("n", "op", "shape", "mode") if k in c} for c in measured["cases"]] == cases


def test_check_63s_file_and_snippet_agree_every_case_parses_each_control_once_showing_its_placeholder_in_the_bare_form():
    body = load(OFFICE_JS / "63-placeholders.docx").body
    regions = [r for p in body.paragraphs if (r := checks.parse_region_63(p.text)) is not None]
    assert regions == CASES_63
    for case in CASES_63:
        if case["shape"].startswith("run"):
            assert len(body.search(f"PLAIN{case['n']}A", **checks.TOKEN)) == 1, case["n"]
    controls = body.get_content_controls()
    assert [c.tag for c in controls] == [f"c63-{c['n']}" for c in CASES_63]
    assert all(c.is_showing_placeholder and c.text == "" for c in controls)
    assert [c.tag for c in controls if c.xml_mapping.is_mapped] == [
        f"c63-{c['n']}" for c in CASES_63 if c["shape"].endswith("-bound")
    ]
    # no w:placeholder: the bare form (CR-019 section 4.1)
    assert {c.placeholder_entry for c in controls} == {None}
    # the file is the one Word read: its paragraphs and controls are Word's save of it as opened
    opened = load(OFFICE_JS / "63-placeholders-word365.docx").body
    assert [p.text for p in opened.paragraphs] == [p.text for p in body.paragraphs]


def test_check_64s_file_and_snippet_agree_every_case_parses_each_control_in_words_form():
    package = load(OFFICE_JS / "64-placeholders-word.docx")
    body = package.body
    regions = [r for p in body.paragraphs if (r := checks.parse_region_64(p.text)) is not None]
    assert regions == CASES_64
    with_control = [c for c in CASES_64 if not c["shape"].startswith("made-")]
    controls = body.get_content_controls()
    assert [c.tag for c in controls] == [f"c64-{c['n']}" for c in with_control]
    assert [c.is_showing_placeholder for c in controls] == [not c["shape"].endswith("-value") for c in with_control]
    assert [c.text for c in controls if c.text] == [
        f"VALUE{c['n']}" for c in with_control if c["shape"].endswith("-value")
    ]
    # each names Word's default entry, which the glossary (check 34's, Word's six parts) holds: Word's form
    assert {c.placeholder_entry for c in controls} == {DEFAULT_ENTRY}
    glossary = package.main_document_part.glossary_document_part
    assert glossary is not None and str(glossary.part_name) == "/word/glossary/document.xml"
    opened = load(OFFICE_JS / "64-opened-word365.docx").body
    assert [p.text for p in opened.paragraphs] == [p.text for p in body.paragraphs]


def test_word_365s_answers_are_core_tss_word_run_through_this_ports_digest():
    run = word("63-placeholders-word365-run.json")
    assert [digest(c) for c in run["cases"]] == WORD_RUN_63
    assert not [c for c in run["cases"] if c.get("error")], "Office JS refused nothing"
    # every bound node left empty by Word's write in the bare form (core-ts section 63.2)
    assert {c["after"]["node"] for c in run["cases"] if c["shape"].endswith("-bound")} == {""}
    run = word("64-placeholders-word-word365-run.json")
    assert [digest(c) for c in run["cases"]] == WORD_RUN_64
    assert all(re.match(r"^\d+ C-- in ", d) for d in WORD_RUN_64)
    # in Word's form every write reached the node (core-ts section 63.5)
    assert [c["after"]["node"] for c in run["cases"] if re.search(r"-(bound|value)$", c["shape"])] == [
        f"NEW{n}" for n in range(43, 53)
    ]


def test_the_report_modes_read_every_region_as_word_365_does():
    for runner, name in (
        (run_63, "63-placeholders-word365-report.json"),
        (run_64, "64-placeholders-word-word365-report.json"),
    ):
        _package, ours = runner("report")
        theirs = word(name)
        assert ours["caseCount"] == theirs["caseCount"]
        assert [view(c) for c in ours["cases"]] == [view(c) for c in theirs["cases"]], name


def test_check_63s_run_is_word_365s_but_for_the_differences_pinned():
    _package, ours = run_63()
    assert not [c for c in ours["cases"] if c.get("error")]
    agreeing = _held(word("63-placeholders-word365-run.json")["cases"], ours["cases"], DIFFERS_63)
    # before CR-019, 8 of the 52 writes: 6, 21, 24, 32, 38, 45, 46 and 52; at the first commit 16
    assert agreeing == [6, 10, 11, 12, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 32, 36, 37, 38, 43, 44, 45, 46, 52]
    assert {why for why, _digest, _rest in DIFFERS_63.values()} == {Q1}


def test_check_64s_run_is_word_365s_but_for_the_controls_insert_content_control_made():
    _package, ours = run_64()
    assert not [c for c in ours["cases"] if c.get("error")]
    agreeing = _held(word("64-placeholders-word-word365-run.json")["cases"], ours["cases"], DIFFERS_64)
    # before CR-019, 20 of the 52 writes
    assert len(agreeing) == 45
    assert {why for why, _digest, _rest in DIFFERS_64.values()} == {MADE}
    assert set(DIFFERS_64) == {c["n"] for c in CASES_64 if c["shape"].startswith("made-")} - {
        15,
        16,
        17,
        18,
        37,
        38,
        39,
    }


def test_d7_every_text_write_into_a_bound_control_writes_its_node_in_both_forms():
    # check 64, Word's form: every node written, a placeholder's empty node as much as a value's, as Word's were
    _package, ours = run_64()
    nodes = {c["n"]: c["after"]["node"] for c in ours["cases"] if "node" in c["after"]}
    assert nodes == {n: f"NEW{n}" for n in range(43, 53)}
    # check 63's bare form: Word left every node empty and lost the writes on reopening (core-ts section 63.2); written
    # here, CR-019 section 4.7's departure (core-ts's D7 (a)), where the control still holds the write
    _package, ours = run_63()
    nodes = {c["n"]: c["after"]["node"] for c in ours["cases"] if "node" in c["after"]}
    assert nodes == {**{n: f"NEW{n}" for n in (21, 22, 23, 24, 25, 47, 48, 49, 50, 51)}, 26: "", 52: ""}


# ---------------------------------------------------------------------------
# Word's saves after the run, region by region
# ---------------------------------------------------------------------------

_PREFIXES = {
    W: "w",
    "http://schemas.microsoft.com/office/word/2010/wordml": "w14",
    "http://schemas.microsoft.com/office/word/2023/wordml/word16du": "w16du",
    "http://www.w3.org/XML/1998/namespace": "xml",
}
#: What a comparison of Word's markup with this package's leaves out: rsids, paraIds, a revision's who and when.
_IGNORED = re.compile(r"^(w:rsid\w*|w14:paraId|w14:textId|w:author|w:date|w16du:dateUtc|xml:space)$")


def _name(clark: str) -> str:
    uri, _, local = clark[1:].partition("}") if clark.startswith("{") else ("", "", clark)
    return f"{_PREFIXES.get(uri, uri)}:{local}" if uri else local


def _canon(element: Any) -> str:
    """An element as markup with its prefixes fixed, ``w:sdtEndPr`` and the ignored attributes left out.

    A revision's or a marker's ``w:id`` is left out; a control's own (in
    ``w:sdtPr``) is kept.
    """
    tag = _name(element.tag)
    if tag == "w:sdtEndPr":
        return ""
    attributes = []
    for key, value in sorted(element.attrib.items()):
        name = _name(key)
        if _IGNORED.match(name) or (name == "w:id" and _name(element.getparent().tag) != "w:sdtPr"):
            continue
        attributes.append(f' {name}="{value}"')
    inner = (element.text or "") + "".join(_canon(child) + (child.tail or "") for child in element)
    return f"<{tag}{''.join(attributes)}>{inner}</{tag}>" if inner else f"<{tag}{''.join(attributes)}/>"


def _regions(document: bytes) -> dict[int, str]:
    """Each case's region in a saved main document part: the blocks between its BEGIN and END paragraphs."""
    body = etree.fromstring(document).find(f"{{{W}}}body")
    out: dict[int, str] = {}
    current: list[Any] | None = None
    n = None
    for child in body:
        text = "".join(child.itertext()) if _name(child.tag) == "w:p" else ""
        begin = re.match(r"^BEGIN(\d+) ", text)
        if begin:
            n, current = int(begin.group(1)), []
        elif n is not None and text == f"END{n}":
            out[n] = "".join(_canon(block) for block in current or [])
            n, current = None, None
        elif current is not None:
            current.append(child)
    return out


def _saved_regions(package: Any) -> dict[int, str]:
    return _regions(zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml"))


def _words_regions(name: str) -> dict[int, str]:
    return _regions(zipfile.ZipFile(OFFICE_JS / name).read("word/document.xml"))


def _plain(region: str) -> str:
    """Word's markup with check 63's look of the text written taken out: its ``PlaceholderText``, ``w:color auto`` and
    ``w:rPrChange``, which checks 58 and 67 contradict (CR-019 Q1)."""

    def strip(found: re.Match[str]) -> str:
        r_pr = re.sub(
            r'<w:rStyle w:val="PlaceholderText"/>|<w:color w:val="auto"/>|<w:rPrChange>[\s\S]*?</w:rPrChange>',
            "",
            found.group(1),
        )
        return f"<w:r>{f'<w:rPr>{r_pr}</w:rPr>' if r_pr else ''}<w:t>NEW"

    return re.sub(r"<w:r><w:rPr>([\s\S]*?)</w:rPr><w:t>NEW", strip, region)


def _deleted_look(region: str) -> str:
    """Word's deleted placeholder runs as they were: in a bound plain-text control Word took the style off them, recorded
    in a ``w:rPrChange`` (check 63 cases 47 to 51), core-ts's D2 (a), which checks 58 and 67 contradict (CR-019 Q1)."""
    return region.replace(
        '<w:rPr><w:rPrChange><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr></w:rPrChange></w:rPr><w:delText>',
        '<w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:delText>',
    )


def _block_level(region: str) -> str:
    """A control Word's save made run-level in its paragraph, put back at block level around it (D6 (b), CR-019 4.5)."""
    return re.sub(
        r"<w:p>((?:<w:pPr>[\s\S]*?</w:pPr>)?)<w:sdt>(<w:sdtPr>[\s\S]*?</w:sdtPr>)<w:sdtContent>([\s\S]*?)</w:sdtContent>"
        r"</w:sdt></w:p>",
        r"<w:sdt>\2<w:sdtContent><w:p>\1\3</w:p></w:sdtContent></w:sdt>",
        region,
    )


def _runs_joined(region: str) -> str:
    """Adjacent runs of plain text as one: Word's save leaves a run boundary where this package's text joined a run."""
    joined = re.sub(
        r"<w:r><w:t>([^<]*)</w:t></w:r><w:r><w:t>([^<]*)</w:t></w:r>", r"<w:r><w:t>\1\2</w:t></w:r>", region
    )
    return joined if joined == region else _runs_joined(joined)


def _no_layout(region: str) -> str:
    """Word's markup without what its layout writes on save (``w:lastRenderedPageBreak``)."""
    return region.replace("<w:lastRenderedPageBreak/>", "")


D6 = "D6 (b), decided in core-ts: the control left at block level, where Word's save made it run-level (CR-019 4.11)"
RUNS = "Word's save leaves the text a run of its own, where this package's joined the run before (CR-019 4.11)"

#: Check 63's regions whose saved markup differs from Word's save after the run: ``n: (why, the forms that make the
#: two alike, applied to both)``; None where the difference is pinned as a difference only.
SAVES_63: dict[int, tuple[str, tuple[Any, ...] | None]] = {
    **{
        n: (Q1, (_plain,)) for n in (1, 2, 3, 4, 5, 7, 8, 9, 13, 14, 15, 16, 27, 28, 29, 30, 31, 33, 34, 35, 39, 40, 41)
    },
    42: (Q1 + "; and Word's layout", (_plain, _no_layout)),
    **{n: (Q1, (_plain, _deleted_look)) for n in (47, 48, 49, 50, 51)},
    **{n: (D6, (_block_level,)) for n in (10, 11, 37)},
    # tracked 'After': the mark marked inserted is the new paragraph's, where Word's, its control run-level, is the
    # mark before the new paragraph, the control's paragraph's
    36: (D6 + "; the inserted mark the new paragraph's", None),
    18: (RUNS, (_runs_joined,)),
}

#: Check 64's regions whose saved markup differs from Word's: the controls insert_content_control made.
SAVES_64: dict[int, tuple[str, tuple[Any, ...] | None]] = {
    n: (MADE, None) for n in (15, 16, 17, 18, 19, 20, 21, 36, 37, 38, 39, 40, 41, 42)
}


def _held_to_save(
    ours: dict[int, str], theirs: dict[int, str], pinned: dict[int, tuple[str, tuple | None]]
) -> list[int]:
    """The regions whose markup is Word's; a pinned one Word's through its forms, or differing as pinned."""
    assert sorted(ours) == sorted(theirs)
    same = []
    for n, markup in theirs.items():
        if n not in pinned:
            assert ours[n] == markup, n
            same.append(n)
            continue
        _why, forms = pinned[n]
        assert ours[n] != markup, f"{n} is Word's now: unpin it"
        if forms is not None:
            mine = ours[n]
            for form in forms:
                mine, markup = form(mine), form(markup)
            assert mine == markup, n
    return same


def test_check_63s_saved_markup_is_words_save_after_the_run_but_for_the_differences_pinned():
    package, _ours = run_63()
    same = _held_to_save(_saved_regions(package), _words_regions("63-run-word365.docx"), SAVES_63)
    # D3's markers among them (12, 19, 20, 26, 38, 44 to 46, 52), and the bare form's tracked deletion's place
    assert same == [6, 12, 17, 19, 20, 21, 22, 23, 24, 25, 26, 32, 38, 43, 44, 45, 46, 52, 53, 54]


def test_check_64s_saved_markup_is_words_save_after_the_run_but_for_the_controls_insert_content_control_made():
    package, _ours = run_64()
    same = _held_to_save(_saved_regions(package), _words_regions("64-run-word365.docx"), SAVES_64)
    # Word's form: every write the placeholder away and the text plain, tracked as an insertion alone
    assert len(same) == 38


# ---------------------------------------------------------------------------
# this package's own: the routes, the ranges and the nodes (CR-019 sections 4.2, 4.4 and 4.7)
# ---------------------------------------------------------------------------

_PLACEHOLDER_RUN = f'<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr><w:t>{PLACEHOLDER}</w:t></w:r>'


def _control_xml(identifier: int, content: str, *, entry: str | None = None, extra: str = "") -> str:
    named = f'<w:placeholder><w:docPart w:val="{entry}"/></w:placeholder>' if entry is not None else ""
    return (
        f'<w:sdt><w:sdtPr><w:id w:val="{identifier}"/>{named}<w:showingPlcHdr/>{extra}</w:sdtPr>'
        f"<w:sdtContent>{content}</w:sdtContent></w:sdt>"
    )


def _block_placeholder(entry: str | None = None, *, tracked: bool = False) -> tuple[Any, Any]:
    package = create_package()
    package.body.insert_xml(_control_xml(1, f"<w:p>{_PLACEHOLDER_RUN}</w:p>", entry=entry), location="End")
    if tracked:
        package.change_tracking_mode = "TrackAll"
    return package, package.body.content_controls[0]


@pytest.mark.parametrize("entry", [None, DEFAULT_ENTRY])
@pytest.mark.parametrize("tracked", [False, True])
@pytest.mark.parametrize(
    "route",
    ["control.insert_text", "content.insert_text", "paragraph.insert_text", "paragraph.text", "paragraph.splice"],
)
def test_every_route_into_a_block_placeholder_writes_one_paragraph_in_its_place_the_flag_cleared(route, tracked, entry):
    package, control = _block_placeholder(entry, tracked=tracked)
    paragraph = control.paragraphs[0]
    if route == "control.insert_text":
        control.insert_text("VALUE", location="End")
    elif route == "content.insert_text":
        control.get_range("Content").insert_text("VALUE", location="End")
    elif route == "paragraph.insert_text":
        paragraph.insert_text("VALUE", location="Start")
    elif route == "paragraph.text":
        paragraph.text = "VALUE"
    else:
        paragraph.splice(0, 0, "VALUE")
    xml = to_xml(control.element)
    assert (control.text, control.is_showing_placeholder, len(control.paragraphs)) == ("VALUE", False, 1), xml
    # what the control shows, its deleted runs aside
    shown = re.sub(r"<w:del [\s\S]*?</w:del>", "", xml)
    assert "PlaceholderText" not in shown and PLACEHOLDER not in shown, xml
    # the bare form's text takes the placeholder run's look but the style; Word's form's is plain (check 64)
    assert ("<w:b/>" in shown) == (entry is None), xml
    # tracked, the insertion first; in the bare form the placeholder's runs then, deleted (check 63), in Word's form not
    # at all (check 64)
    assert ("<w:ins " in xml) == tracked and ("<w:del " in xml) == (tracked and entry is None), xml
    if tracked and entry is None:
        assert xml.index("<w:ins ") < xml.index("<w:del ") and f"<w:delText>{PLACEHOLDER}</w:delText>" in xml
    again = load(io.BytesIO(package.save()), options=None).body.content_controls[0]
    assert (again.text, again.is_showing_placeholder) == ("VALUE", False)


def test_the_bare_forms_insert_paragraph_at_end_writes_in_the_placeholders_place_and_leaves_the_flag_on_check_63_case_3():
    for tracked in (False, True):
        _package, control = _block_placeholder(None, tracked=tracked)
        made = control.insert_paragraph("NEW", location="End")
        assert len(control.paragraphs) == 1 and made.element is control.paragraphs[0].element
        assert control.is_showing_placeholder and control.text == ""
        xml = to_xml(control.element)
        assert ">NEW</w:t>" in xml and ("<w:delText>" in xml) == tracked


@pytest.mark.parametrize("tracked", [False, True])
def test_a_controls_own_content_range_deleted_keeps_it_and_its_flag_its_placeholder_going_in_the_bare_form(tracked):
    # core-ts check 63 cases 6, 17, 32 and 43; in Word's form nothing changes (unmeasured, CR-019 section 4.10)
    for entry in (None, DEFAULT_ENTRY):
        package, control = _block_placeholder(entry, tracked=tracked)
        before = to_xml(control.element)
        control.get_range("Content").delete()
        xml = to_xml(control.element)
        assert len(package.body.content_controls) == 1 and control.is_showing_placeholder
        if entry is not None:
            assert xml == before
        else:
            assert (f"<w:delText>{PLACEHOLDER}</w:delText>" in xml) == tracked
            assert (f"<w:t>{PLACEHOLDER}</w:t>" in xml) is False


def test_a_run_level_controls_whole_range_deleted_takes_it_as_its_replace_does():
    # D3, unmeasured for the delete (CR-019 section 4.10): untracked the control goes; tracked it is marked deleted
    package, control = _run_control(_PLACEHOLDER_RUN, showing=True)
    control.get_range("Whole").delete()
    assert (package.body.paragraphs[0].text, package.body.content_controls) == ("A  B", [])
    package, control = _run_control(_PLACEHOLDER_RUN, showing=True)
    package.change_tracking_mode = "TrackAll"
    control.get_range("Whole").delete()
    xml = to_xml(package.body.paragraphs[0].element)
    assert re.search(
        r"<w:customXmlDelRangeStart [^>]*/><w:sdt>.*<w:sdtContent><w:customXmlDelRangeEnd [^>]*/><w:del ", xml
    )


def test_words_form_insert_paragraph_at_end_empties_the_placeholders_paragraph_and_puts_the_new_one_after_check_64_case_6():
    for tracked in (False, True):
        _package, control = _block_placeholder(DEFAULT_ENTRY, tracked=tracked)
        made = control.insert_paragraph("NEW", location="End")
        assert [p.text for p in control.paragraphs] == ["", "NEW"]
        assert made.text == "NEW" and not control.is_showing_placeholder
        assert PLACEHOLDER not in to_xml(control.element)


def _run_control(content: str, *, showing: bool = False) -> tuple[Any, Any]:
    package = create_package()
    flag = "<w:showingPlcHdr/>" if showing else ""
    package.body.insert_xml(
        f'<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="2"/>{flag}</w:sdtPr>'
        f'<w:sdtContent>{content}</w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> B</w:t></w:r></w:p>',
        location="End",
    )
    return package, package.body.content_controls[0]


@pytest.mark.parametrize("tracked", [False, True])
def test_d4_an_empty_run_level_controls_content_range_takes_the_text_inside_it_where_it_stands(tracked):
    package, control = _run_control("")
    # it stands after "A ", where the walk passes it; it was found at 0
    assert (control.get_range("Start").start, control.get_range("Content").text) == (2, "")
    if tracked:
        package.change_tracking_mode = "TrackAll"
    returned = control.get_range("Content").insert_text("X", location="Replace")
    assert (returned.text, control.text, package.body.paragraphs[0].text) == ("X", "X", "A X B")
    inner = to_xml(control.element)
    assert ("<w:ins " in inner) == tracked and ">X</w:t>" in inner


@pytest.mark.parametrize("tracked", [False, True])
def test_q4_a_run_level_controls_content_range_at_start_writes_inside_it_before_its_text(tracked):
    package, control = _run_control("<w:r><w:t>VAL</w:t></w:r>")
    if tracked:
        package.change_tracking_mode = "TrackAll"
    control.get_range("Content").insert_text("X", location="Start")
    control.get_range("Content").insert_text("Y", location="End")
    assert (control.text, package.body.paragraphs[0].text) == ("XVALY", "A XVALY B")
    if tracked:
        assert len(re.findall("<w:ins ", to_xml(control.element))) == 2


def test_a_range_made_from_a_controls_own_range_is_an_ordinary_one():
    _package, control = _run_control("", showing=False)
    own = control.get_range("Content")
    assert own.of_control == (control, "Content")
    assert own.get_range("Start").of_control is None and own.expand_to(own).of_control is None


def _bound_package(content: str, *, block: bool) -> tuple[Any, Any, Any]:
    package = create_package()
    binding = (
        '<w:dataBinding w:prefixMappings="xmlns:ns0=\'http://example.com/g\'" w:xpath="/ns0:greeting[1]/ns0:to[1]" '
        'w:storeItemID="{11111111-2222-4333-8444-555555555555}"/><w:text/>'
    )
    sdt = _control_xml(3, content, extra=binding)
    package.body.insert_xml(
        sdt if block else f'<w:p><w:r><w:t xml:space="preserve">To: </w:t></w:r>{sdt}</w:p>', location="End"
    )
    part = package.custom_xml_parts.add('<greeting xmlns="http://example.com/g"><to/></greeting>')
    control = package.body.get_content_controls()[0]
    control.xml_mapping.set_mapping("/ns0:greeting[1]/ns0:to[1]", "xmlns:ns0='http://example.com/g'", part)
    return package, control, part


@pytest.mark.parametrize("route", ["control.insert_text", "content.insert_text", "paragraph.insert_text"])
@pytest.mark.parametrize("block", [True, False])
def test_d7_a_bound_controls_node_takes_every_routes_write_block_or_run(block, route):
    package, control, part = _bound_package(
        f"<w:p>{_PLACEHOLDER_RUN}</w:p>" if block else _PLACEHOLDER_RUN, block=block
    )
    if route == "control.insert_text":
        control.insert_text("VALUE", location="Replace")
    elif route == "content.insert_text":
        control.get_range("Content").insert_text("VALUE", location="Replace")
    elif block:
        control.paragraphs[0].insert_text("VALUE", location="Replace")
    else:
        # a paragraph's write into the text of a run-level bound control holding a value, at its end
        control.insert_text("VAL", location="Replace")
        package.body.paragraphs[0].splice(len("To: VAL"), len("To: VAL"), "UE")
    assert part.select_single_node("/ns0:greeting[1]/ns0:to[1]").text == "VALUE"
    assert control.text == "VALUE"


# ---------------------------------------------------------------------------
# core-ts's tests that are no replay (placeholders.test.mjs at 020c039): D3 and D6
# ---------------------------------------------------------------------------

_MARKED = (
    '<w:customXmlDelRangeStart/><w:sdt><w:sdtPr><w:alias w:val="Check 63 control {n}"/><w:tag w:val="c63-{n}"/>'
    '<w:id w:val="{id}"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:customXmlDelRangeEnd/><w:del><w:r><w:rPr>'
    '<w:rStyle w:val="PlaceholderText"/></w:rPr><w:delText>Click or tap here to enter text.</w:delText></w:r></w:del>'
    "<w:customXmlDelRangeStart/></w:sdtContent></w:sdt><w:customXmlDelRangeEnd/>"
)


def test_d3_a_replace_or_delete_over_a_whole_control_showing_its_placeholder_takes_it_tracked_in_words_markers_the_insertion_first():
    package, ours = run_63()
    assert not [c for c in ours["cases"] if c.get("error")]
    saved = package.save()
    regions = _regions(zipfile.ZipFile(io.BytesIO(saved)).read("word/document.xml"))

    def marked(n: int) -> str:
        return _MARKED.format(n=n, id=63000 + n)

    # untracked the control goes: the paragraph's Replace (19), its Content delete (20), the control's Whole Replace (18)
    assert regions[19] == "<w:p><w:r><w:t>NEW19</w:t></w:r></w:p>"
    assert regions[20] == "<w:p/>"
    assert regions[18] == "<w:p><w:r><w:t>PLAIN18A NEW18</w:t></w:r><w:r><w:t> PLAIN18B</w:t></w:r></w:p>"
    assert regions[12] == "<w:p/>", "a block control goes, its paragraph emptied in its place"
    # tracked (44 to 46): Word's markers around the control and inside it, its placeholder deleted, a Replace's
    # insertion first
    assert regions[44] == (
        "<w:p><w:r><w:t>PLAIN44A </w:t></w:r><w:ins><w:r><w:t>NEW44</w:t></w:r></w:ins>"
        f"{marked(44)}<w:r><w:t> PLAIN44B</w:t></w:r></w:p>"
    )
    assert regions[45] == (
        "<w:p><w:ins><w:r><w:t>NEW45</w:t></w:r></w:ins><w:del><w:r><w:delText>PLAIN45A </w:delText></w:r></w:del>"
        f"{marked(45)}<w:del><w:r><w:delText> PLAIN45B</w:delText></w:r></w:del></w:p>"
    )
    assert regions[46] == (
        f"<w:p><w:del><w:r><w:delText>PLAIN46A </w:delText></w:r></w:del>{marked(46)}"
        "<w:del><w:r><w:delText> PLAIN46B</w:delText></w:r></w:del></w:p>"
    )
    assert re.match(
        r"^<w:customXmlDelRangeStart/><w:sdt>.*<w:sdtContent><w:customXmlDelRangeEnd/><w:p><w:del><w:r><w:rPr>"
        r'<w:rStyle w:val="PlaceholderText"/></w:rPr><w:delText>Click or tap here to enter text\.</w:delText></w:r>'
        r"</w:del></w:p><w:customXmlDelRangeStart/></w:sdtContent></w:sdt><w:customXmlDelRangeEnd/>$",
        regions[38],
    )
    # the bare form's tracked write: the insertion, then the placeholder's deletion (28, 47): core-ts's D2 (a) gives the
    # insertion Word's PlaceholderText, w:color auto and w:rPrChange, which checks 58 and 67 contradict (CR-019 Q1)
    placeholder = '<w:del><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:delText>Click or tap here to enter text.</w:delText></w:r></w:del>'
    assert regions[28] == (
        '<w:sdt><w:sdtPr><w:alias w:val="Check 63 control 28"/><w:tag w:val="c63-28"/><w:id w:val="63028"/></w:sdtPr>'
        f"<w:sdtContent><w:p><w:ins><w:r><w:t>NEW28</w:t></w:r></w:ins>{placeholder}</w:p></w:sdtContent></w:sdt>"
    )
    assert (
        f"<w:sdtContent><w:p><w:ins><w:r><w:t>NEW47</w:t></w:r></w:ins>{placeholder}</w:p></w:sdtContent>"
        in regions[47]
    )
    # accepted, the controls taken go and the text stays; rejected, they come back showing their placeholder
    from docx4j_py.model import revisions

    accepted = load(io.BytesIO(saved))
    revisions.accept_all(accepted)
    after = _saved_regions(accepted)
    assert after[45] == "<w:p><w:r><w:t>NEW45</w:t></w:r></w:p>"
    assert _runs_joined(after[44]) == "<w:p><w:r><w:t>PLAIN44A NEW44 PLAIN44B</w:t></w:r></w:p>"
    rejected = load(io.BytesIO(saved))
    revisions.reject_all(rejected)
    restored = checks.control_by_tag(rejected.body, "c63-45")
    assert [restored.is_showing_placeholder, restored.text, restored.paragraphs[0].text] == [
        True,
        "",
        "PLAIN45A  PLAIN45B",
    ]


def test_d6_b_a_paragraph_put_beside_the_only_paragraph_of_a_block_control_showing_its_placeholder_goes_outside_it_at_block_level():
    package, _ours = run_63()
    regions = _saved_regions(package)
    # Word's save has the control run-level in its paragraph (core-ts section 63.2); here it stays as it was
    held = (
        '<w:sdtContent><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>'
        "<w:t>Click or tap here to enter text.</w:t></w:r></w:p></w:sdtContent></w:sdt>"
    )
    for n, before in ((10, False), (11, True), (36, False), (37, True)):
        region = regions[n]
        assert held in region, n
        assert re.match(
            rf"^<w:p>.*NEW{n}.*</w:p><w:sdt>" if before else rf"^<w:sdt>.*</w:sdt><w:p>.*NEW{n}.*</w:p>$", region
        ), n
