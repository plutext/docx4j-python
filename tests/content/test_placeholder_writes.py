"""Writing into a content control showing its placeholder: CR-019 Phases A and A2, the port of core-ts's CR-002 section 63.

    .venv-fork/bin/python -m pytest tests/content/test_placeholder_writes.py

docx4j-core-ts ran Word 365 checks 63 and 64 on what Word does with every text
write into a content control showing its placeholder (its CR-002 sections 63.2
and 63.5; Word 365 16.0.20326.20232, 2026-10-24), and implemented the answers at
``020c039``. Check 63's first run proved to have been made in Developer > Design
Mode; core-ts's check 67 (``2bafbcb``) and check 63's run 2, with Design Mode off
(its sections 63.13 to 63.16), gave the bare form Word's form's answers, and
core-ts implemented them at ``d84f3da`` (its section 63.17): CR-019 Phase A2.
core-ts's ``test/placeholders.test.mjs`` replays them through its Office JS shim;
this package has none, so ``tests/office_js/checks.py`` ports the Script Lab
snippets step for step (:func:`run_check_63`, :func:`run_check_64`,
:func:`run_check_67`), and this file holds this package's run to Word's, case
by case, through core-ts's ``digest`` and ``view``, then the markup of each
region against Word's save after the same run (rsids, paraIds, ``w:sdtEndPr``
and revision ``w:id`` values left out, as CR-016's check 58 comparison leaves
them).

* **Check 63** is the bare form (no ``w:placeholder``, no glossary): what
  docx4j's binding and this package write. Its run 2 is the oracle; run 1, in
  Design Mode, is kept as the record of what Design Mode did.
* **Check 64** is Word's form (``w:placeholder`` naming Word's default entry,
  check 34's glossary), with controls Word makes and bound controls' nodes.
* **Check 67** is check 58's three bare-form writes in three files, four phases
  and through two lookups: all 72 a plain run.

What still differs is pinned, with its reason: the controls
``insert_content_control`` makes (Phase B's, D8), and in Word's saves D6 (b)'s
level and Word's run boundaries. Then core-ts's tests that are no replay: D3's
markup, accepted and rejected; D6's level; D7's nodes in both forms; D10's
filler. And this package's own: every route, the control's own ranges, Q4, the
nodes, the filler's readings.
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

#: The files, as core-ts 020c039 and d84f3da have them: each check's file (git-ignored there, written by its
#: build.mjs before Word's run, and read here as Word's save of it as opened reads), Word's JSONs and Word's saves
#: (its manifest.json's SHA-1s). Check 63's run 2 (d84f3da) is the same file and snippet with Design Mode off.
SHA1 = {
    "63-placeholders.docx": "788396d96be54ac1b4f825d8c1de5d1eb9fb93c3",
    "63-placeholders-word365-report.json": "a84e8987b3d79653ea91136ef1a120e61e401418",
    "63-placeholders-word365-run.json": "8f0404a4b301669934d13315e295e3dc44e0e284",
    "63-placeholders-word365.docx": "b8dc3e3b7b272a7c7bf0dbf7fdb76336ac78a3e8",
    "63-run-word365.docx": "6c118dcce476b37e7d3f61ed252af0eec6c325b2",
    "63-placeholders-word365-run2-report.json": "382acc8bc43f874925a700706de8485ef72ab742",
    "63-placeholders-word365-run2-run.json": "a5852075422ccf9b927cf9cbf29c0629aaa9176e",
    "63-run2-opened-word365.docx": "1c567c54f627ccb7a1c8e7fb50ec16bff98b5e05",
    "63-run2-word365.docx": "9ec9fdb8e6591ab359c30f112e106a11892d2383",
    "64-placeholders-word.docx": "fc2e15bfb4c50e1111cfb834cb132e8f36fc8450",
    "64-placeholders-word-word365-report.json": "eb265f08a79c29f71d658e561af7d60eca378fbb",
    "64-placeholders-word-word365-run.json": "4da651a179bfbf0cc06df6c9d20861f30bff7269",
    "64-opened-word365.docx": "bff704e268525a444db9732af7cd44df6d1404a0",
    "64-run-word365.docx": "62220a054f033c842ed4ba1f7711c9e0f314a361",
}

#: Check 67's files, as core-ts 2bafbcb has them (unchanged at d84f3da): its three files (git-ignored there, written
#: by build.mjs on 2026-10-25 before Word's run), Word's three JSONs, each holding all five steps, and Word's six
#: saves, which its test/check67/manifest.json lists with these SHA-1s.
SHA1_67 = {
    "67a-plain.docx": "4a5bc6e08b57f9bd0dd058b6133dca9a86468083",
    "67b-custom-xml.docx": "786a32fdfaf609bb8b8eebd319e3a8fe6a33f7de",
    "67c-revisions.docx": "c8ec3522455542da300cf07a98e8b6358dadc514",
    "67a-plain-word365.json": "420656b467ec0edfb5caf80b5ec21017d332f262",
    "67b-custom-xml-word365.json": "1adcda39bcab5a6cbe22ae221f702736a1ef83e1",
    "67c-revisions-word365.json": "80dc2271ea4ebb38e385c2aec1442a7a3509ac1a",
    "67a-opened-word365.docx": "7c8e453d275f5dda1f656e1179b5ac0e05b4550b",
    "67a-run-word365.docx": "747727de000195fee7c830215e2c2efe19d23dcc",
    "67b-opened-word365.docx": "f3cbb3400387c07c48fe4a592bc1a29c0b223e55",
    "67b-run-word365.docx": "e79485707dfcc6d0bc233a82e39d8f3970641dfe",
    "67c-opened-word365.docx": "21caf9a02ca26a2a4aa5725d0a5165c25b4b38e3",
    "67c-run-word365.docx": "38f291e350c76b63a469ddab89b25985fbf4f75e",
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
    (``s`` where only deleted runs have it). Where ``NEW<n>`` (``FILL<n>`` in
    check 67, check 58's word) landed: ``in``, ``out`` (outside the control), or
    ``-``. Then the region's paragraphs' text, ``|`` between, and the write's
    error code if it threw. core-ts's at ``d84f3da``.
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
    tokens = (f"NEW{case['n']}", f"FILL{case['n']}")
    where = "in" if any(t in content for t in tokens) else "out" if any(t in xml for t in tokens) else "-"
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


#: Word 365's run 1 of check 63, in Design Mode, case by case, as `digest` reads it: core-ts's `WORD_RUN` (its CR-002
#: section 63.2). Kept as the record of what Design Mode did; the oracle is run 2.
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


#: Word 365's run 2 of check 63, Design Mode off (2026-10-25), as `digest` reads it: core-ts's `WORD_RUN2` (its CR-002
#: section 63.16). The bare form gives Word's form's answers: every write plain and the flag cleared, only the
#: insertion tracked, insert_paragraph clearing the flag (3, 29); a Content delete keeps the flag over Word's filler
#: (6, 17, 32, 43); a whole control still taken (12, 18 to 20, 26, 38, 44 to 46, 52), tracked without its flag.
WORD_RUN2_63 = [
    "1 C-- in NEW1",
    "2 C-- in NEW2",
    "3 C-- in |NEW3",
    "4 C-- in NEW4",
    "5 C-- in NEW5",
    "6 CP- - ",
    "7 C-- in NEW7",
    "8 C-- in NEW8",
    "9 C-- in NEW9",
    "10 CPS out |NEW10",
    "11 CPS out NEW11|",
    "12 x-- - ",
    "13 C-- in PLAIN13A NEW13 PLAIN13B",
    "14 C-- in PLAIN14A NEW14 PLAIN14B",
    "15 C-- in PLAIN15A NEW15 PLAIN15B",
    "16 C-- in PLAIN16A NEW16 PLAIN16B",
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
    "27 C-- in NEW27",
    "28 C-- in NEW28",
    "29 C-- in |NEW29",
    "30 C-- in NEW30",
    "31 C-- in NEW31",
    "32 CP- - ",
    "33 C-- in NEW33",
    "34 C-- in NEW34",
    "35 C-- in NEW35",
    "36 CPS out |NEW36",
    "37 CPS out NEW37|",
    "38 C-- - ",
    "39 C-- in PLAIN39A NEW39 PLAIN39B",
    "40 C-- in PLAIN40A NEW40 PLAIN40B",
    "41 C-- in PLAIN41A NEW41 PLAIN41B",
    "42 C-- in PLAIN42A NEW42 PLAIN42B",
    "43 CP- - PLAIN43A  PLAIN43B",
    "44 C-- out PLAIN44A NEW44 PLAIN44B",
    "45 C-- out NEW45",
    "46 C-- - ",
    "47 C-- in NEW47",
    "48 C-- in NEW48",
    "49 C-- in NEW49",
    "50 C-- in PLAIN50A NEW50 PLAIN50B",
    "51 C-- in PLAIN51A NEW51 PLAIN51B",
    "52 C-- out NEW52",
    "53 CP- - PLAIN53A  PLAIN53B",
    "54 CP- - ",
]

#: The writes in which Design Mode's run 1 differs from run 2 (core-ts's 31, section 63.17): the style kept,
#: insert_paragraph's flag left on, the placeholder marked deleted under tracking, a Content delete leaving no filler.
RUN_1_DIFFERS = [
    *(1, 2, 3, 4, 5, 7, 8, 9, 13, 14, 15, 16),
    *(27, 28, 29, 30, 31, 32, 33, 34, 35, 38, 39, 40, 41, 42, 43, 44, 45, 46, 52),
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
MADE = "a control insert_content_control made, which shows no placeholder here: Phase B's (CR-019 section 4.8, D8)"

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
        ("63-placeholders-word365-run2", 63, 54, CASES_63),
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
    # the file is the one Word read, twice: its paragraphs and controls are Word's saves of it as opened, run 1's and
    # run 2's
    for name in ("63-placeholders-word365.docx", "63-run2-opened-word365.docx"):
        opened = load(OFFICE_JS / name).body
        assert [p.text for p in opened.paragraphs] == [p.text for p in body.paragraphs], name
        assert [(c.tag, c.is_showing_placeholder, c.text) for c in opened.get_content_controls()] == [
            (c.tag, c.is_showing_placeholder, c.text) for c in controls
        ], name


def _settings(name: str) -> str:
    return zipfile.ZipFile(OFFICE_JS / name).read("word/settings.xml").decode("utf-8")


def test_check_63s_run_1_was_made_in_design_mode_and_run_2_with_it_off():
    # core-ts's "check 63's run 2: Word's two JSONs are of this file's cases, run with Design Mode off": run 1's saves
    # carry w:formsDesign, Design Mode's, which the file didn't hold; run 2's don't
    designed = [
        "<w:formsDesign/>" in _settings(name)
        for name in (
            "63-placeholders-word365.docx",
            "63-run-word365.docx",
            "63-run2-opened-word365.docx",
            "63-run2-word365.docx",
        )
    ]
    assert designed == [True, True, False, False]
    assert "formsDesign" not in _settings("63-placeholders.docx")


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
    # run 1, in Design Mode: every bound node left empty by Word's write in the bare form (core-ts section 63.2)
    assert {c["after"]["node"] for c in run["cases"] if c["shape"].endswith("-bound")} == {""}
    run = word("63-placeholders-word365-run2-run.json")
    assert [digest(c) for c in run["cases"]] == WORD_RUN2_63
    run = word("64-placeholders-word-word365-run.json")
    assert [digest(c) for c in run["cases"]] == WORD_RUN_64
    assert all(re.match(r"^\d+ C-- in ", d) for d in WORD_RUN_64)
    # in Word's form every write reached the node (core-ts section 63.5)
    assert [c["after"]["node"] for c in run["cases"] if re.search(r"-(bound|value)$", c["shape"])] == [
        f"NEW{n}" for n in range(43, 53)
    ]


def test_word_365s_answers_with_design_mode_off_check_63_run_2_the_bare_form_gives_words_forms_answers():
    """core-ts: "Word 365's answers with Design Mode off (check 63, run 2): the bare form gives Word's form's answers; a
    Replace or delete over the whole control still takes it" (its section 63.16)."""
    report, run = word("63-placeholders-word365-run2-report.json"), word("63-placeholders-word365-run2-run.json")
    # opened, Word reads the file as run 1 did
    assert [digest(c) for c in report["cases"]] == [
        digest(c) for c in word("63-placeholders-word365-report.json")["cases"]
    ]
    assert not [c for c in run["cases"] if c.get("error")], "Office JS refused nothing"
    cases = run["cases"]
    # a write into the control: plain, as in Word's form (check 64), where run 1 kept the style and added w:color auto
    for n in (1, 2, 4, 5, 7, 8, 9, 13, 14, 15, 16, 27, 28, 30, 31, 33, 34, 35, 39, 40, 41, 42):
        assert not re.search(r'PlaceholderText|w:color w:val="auto"', cases[n - 1]["after"]["ooxml"]), n
    # tracked, only the insertion is recorded: the placeholder goes unmarked, as in Word's form
    for n in (27, 28, 29, 30, 31, 33, 34, 35, 39, 40, 41, 42, 47, 48, 49, 50, 51):
        ooxml = cases[n - 1]["after"]["ooxml"]
        assert "<w:ins " in ooxml and not re.search(r"<w:del |<w:delText", ooxml), n
    # the control's insertParagraph clears the flag: the placeholder's paragraph left empty, the new one after it (3)
    assert re.search(
        r"</w:sdtPr><w:sdtContent><w:p/><w:p><w:r><w:t>NEW3</w:t></w:r></w:p></w:sdtContent>",
        cases[2]["after"]["ooxml"],
    )
    # a control's Content deleted: the flag stays, over a run of five spaces, untracked and tracked alike (6, 17, 32, 43)
    for n in (6, 17, 32, 43):
        assert re.search(
            r'<w:showingPlcHdr/></w:sdtPr><w:sdtContent>(<w:p>)?<w:r><w:t xml:space="preserve"> {5}</w:t></w:r>'
            r"(</w:p>)?</w:sdtContent>",
            cases[n - 1]["after"]["ooxml"],
        ), n
    # a Replace or delete over the whole control takes it; tracked, Word's markers around it, nothing inside marked
    # deleted and the flag gone (38, 44 to 46, 52)
    for n in (38, 44, 45, 46, 52):
        found = re.search(
            r"<w:customXmlDelRangeStart [^>]*/><w:sdt>([\s\S]*?)</w:sdt><w:customXmlDelRangeEnd [^>]*/>",
            cases[n - 1]["after"]["ooxml"],
        )
        assert found and not re.search(r"showingPlcHdr|<w:del |PlaceholderText", found.group(1)), n
    # every bound control's node holds the write (21 to 25, 47 to 51), as in Word's form, and the controls showed them on
    # reopening (core-ts section 63.16); where the write took the control (26, 52) the node stays empty
    bound = (21, 22, 23, 24, 25, 47, 48, 49, 50, 51)
    assert {n: cases[n - 1]["after"]["node"] for n in (*bound, 26, 52)} == {
        **{n: f"NEW{n}" for n in bound},
        26: "",
        52: "",
    }
    saved = zipfile.ZipFile(OFFICE_JS / "63-run2-word365.docx").read("customXml/item1.xml").decode("utf-8")
    for n in bound:
        assert f"<value{n}>NEW{n}</value{n}>" in saved, n
    # a paragraph beside a block control's only paragraph goes outside it, the control saved run-level, as in run 1
    assert re.match(
        r"^<w:p>[\s\S]*?<w:p><w:sdt>[\s\S]*</w:sdt></w:p><w:p><w:r><w:t>NEW10</w:t></w:r></w:p>",
        cases[9]["after"]["ooxml"],
    )


def test_the_report_modes_read_every_region_as_word_365_does():
    for runner, name in (
        (run_63, "63-placeholders-word365-report.json"),
        (run_63, "63-placeholders-word365-run2-report.json"),
        (run_64, "64-placeholders-word-word365-report.json"),
    ):
        _package, ours = runner("report")
        theirs = word(name)
        assert ours["caseCount"] == theirs["caseCount"]
        assert [view(c) for c in ours["cases"]] == [view(c) for c in theirs["cases"]], name


def test_check_63s_run_is_word_365s_run_2_in_every_write():
    """core-ts: "the shim gives check 63's run 2 in every write (section 63.12 followed run 1, and differed in 31)"."""
    _package, ours = run_63()
    assert not [c for c in ours["cases"] if c.get("error")]
    agreeing = _held(word("63-placeholders-word365-run2-run.json")["cases"], ours["cases"], {})
    # at Phase A's code (67db0d5) 27 of the 52 writes agreed with run 2 (and 23 with run 1)
    assert len(agreeing) == 52
    # and the markup where it tells: plain runs, only insertions tracked, nothing marked inside a taken control
    for n in (1, 2, 13, 14, 27, 28, 39, 40):
        assert not re.search(r'PlaceholderText|w:color w:val="auto"|<w:del ', ours["cases"][n - 1]["after"]["ooxml"]), n


def test_check_63s_run_differs_from_run_1_in_design_modes_31_writes():
    """core-ts: "the shim against check 63's run 1, in Design Mode: it reads the file as Word did, and differs in the
    writes Design Mode made": the record of what Design Mode did, the writes Phase A followed and A2 corrected."""
    _package, ours = run_63()
    run_1 = word("63-placeholders-word365-run.json")["cases"]
    differing = [c["n"] for c, mine in zip(run_1, ours["cases"], strict=True) if view(c) != view(mine)]
    assert differing == RUN_1_DIFFERS
    assert len(differing) == 31
    # the same 31 by which Word's two runs differ
    run_2 = word("63-placeholders-word365-run2-run.json")["cases"]
    assert [c["n"] for c, other in zip(run_1, run_2, strict=True) if view(c) != view(other)] == RUN_1_DIFFERS


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
    # check 63's bare form: every node written, as Word's run 2 wrote them (run 1, in Design Mode, left every node
    # empty, and Word lost the writes on reopening): D7 is Word's behaviour in both forms (CR-019 section 11)
    _package, ours = run_63()
    nodes = {c["n"]: c["after"]["node"] for c in ours["cases"] if "node" in c["after"]}
    assert nodes == {**{n: f"NEW{n}" for n in (21, 22, 23, 24, 25, 47, 48, 49, 50, 51)}, 26: "", 52: ""}
    measured = word("63-placeholders-word365-run2-run.json")["cases"]
    assert nodes == {c["n"]: c["after"]["node"] for c in measured if "node" in c["after"]}


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


D6 = "D6 (b), decided in core-ts: the control left at block level, where Word's save made it run-level (CR-019 4.11)"
RUNS = "Word's save leaves the text a run of its own, where this package's joined the run before (CR-019 4.11)"

#: Check 63's regions whose saved markup differs from Word's save after run 2: ``n: (why, the forms that make the
#: two alike, applied to both)``; None where the difference is pinned as a difference only. Every other region is
#: Word's, D3's markers, D10's filler and the tracked insertions among them.
SAVES_63: dict[int, tuple[str, tuple[Any, ...] | None]] = {
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


def test_check_63s_saved_markup_is_words_save_after_run_2_but_for_d6_b_and_a_run_boundary():
    package, _ours = run_63()
    same = _held_to_save(_saved_regions(package), _words_regions("63-run2-word365.docx"), SAVES_63)
    # at Phase A's code (67db0d5) 22 of the 54 regions were run 2's (20 run 1's); D3's markers among them now (12, 19,
    # 20, 26, 38, 44 to 46, 52), D10's filler (6, 17, 32, 43) and insert_paragraph's emptied paragraph (3, 29)
    assert len(same) == 49
    assert sorted(set(range(1, 55)) - set(same)) == sorted(SAVES_63)


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
    # the bare form's text takes the placeholder run's own look but the style, CR-016 Phase B's (a plain run for every
    # placeholder Word measured, whose run had the style alone); Word's form's is plain (check 64). core-ts at d84f3da
    # writes plain in both forms, the bold dropped too: unmeasured, CR-019 section 11.4
    assert ("<w:b/>" in shown) == (entry is None), xml
    # tracked, only the insertion is recorded, in either form: the placeholder goes unmarked (check 64, check 63's run 2;
    # core-ts's D5 at d84f3da), where Phase A put the bare form's placeholder back in a w:del (run 1, Design Mode's)
    assert ("<w:ins " in xml) == tracked and "<w:del " not in xml and "<w:delText>" not in xml, xml
    again = load(io.BytesIO(package.save()), options=None).body.content_controls[0]
    assert (again.text, again.is_showing_placeholder) == ("VALUE", False)


@pytest.mark.parametrize("entry", [None, DEFAULT_ENTRY])
def test_insert_paragraph_at_end_empties_the_placeholders_paragraph_and_puts_the_new_one_after_in_either_form(entry):
    # check 64 case 6 (Word's form), check 63's run 2 case 3 (the bare form; core-ts's D1 at d84f3da), where Phase A's
    # bare form wrote the text in the placeholder's place and left the flag on (run 1, Design Mode's)
    for tracked in (False, True):
        _package, control = _block_placeholder(entry, tracked=tracked)
        made = control.insert_paragraph("NEW", location="End")
        assert [p.text for p in control.paragraphs] == ["", "NEW"]
        assert made.text == "NEW" and not control.is_showing_placeholder
        xml = to_xml(control.element)
        assert PLACEHOLDER not in xml and "<w:del " not in xml and ("<w:ins " in xml) == tracked


def _filler(control: Any) -> list[tuple[str, str | None]]:
    """A control's runs as ``(text, their w:rPr's markup or None)``: what Word's filler is made of."""
    out = []
    for run in control._content_runs():
        text = "".join(getattr(item, "value", "") or "" for item in run.content)
        out.append((text, to_xml(run.r_pr) if run.r_pr is not None else None))
    return out


@pytest.mark.parametrize("tracked", [False, True])
@pytest.mark.parametrize("form", ["block", "run"])
def test_d10_a_controls_own_content_range_deleted_keeps_it_and_its_flag_over_words_filler_in_the_bare_form(
    form, tracked
):
    # check 63's run 2 cases 6, 17, 32 and 43 (core-ts's D10 (a), showFiller): five spaces in the control's own run
    # properties, the flag on, nothing marked under tracking; Phase A left the content empty (run 1), and a run-level
    # control showing its placeholder over an empty w:sdtContent is what Word refused to leave Design Mode over
    if form == "block":
        package, control = _block_placeholder(None, tracked=tracked)
    else:
        package, control = _run_control(_PLACEHOLDER_RUN, showing=True, run_properties="<w:rPr><w:i/></w:rPr>")
        if tracked:
            package.change_tracking_mode = "TrackAll"
    control.get_range("Content").delete()
    xml = to_xml(control.element)
    assert len(package.body.content_controls) == 1 and control.is_showing_placeholder, xml
    assert _filler(control) == [("     ", None if form == "block" else to_xml(control.run_properties))], xml
    assert "<w:del " not in xml and "<w:ins " not in xml and PLACEHOLDER not in xml, xml
    # the filler reads as no text to every Office JS reader while the flag is on (CR-016 Phase B)
    assert control.text == "" and package.body.search("     ", match_case=True) == []
    assert package.body.text == ("" if form == "block" else "A  B")
    # the text model reads it as it reads any placeholder: a placeholder run, left out with placeholders=False
    from docx4j_py.model.text import TextOptions, extract, plain_text

    assert "     " in plain_text(extract(package))
    assert plain_text(extract(package, TextOptions(placeholders=False))).strip() == ("" if form == "block" else "A  B")
    again = load(io.BytesIO(package.save())).body.content_controls[0]
    assert (again.text, again.is_showing_placeholder, _filler(again)) == ("", True, _filler(control))


def test_d10_in_words_form_a_controls_content_range_deleted_changes_nothing():
    # unmeasured (CR-019 section 4.10): the placeholder is no content in Word's form
    for tracked in (False, True):
        _package, control = _block_placeholder(DEFAULT_ENTRY, tracked=tracked)
        before = to_xml(control.element)
        control.get_range("Content").delete()
        assert to_xml(control.element) == before


@pytest.mark.parametrize("form", ["block", "run"])
def test_d10_a_control_holding_text_emptied_by_its_content_range_untracked_shows_words_filler_in_the_bare_form(form):
    # core-ts's showEmptied at d84f3da (check 52 case 13): the filler once the control holds no text; tracked, the text
    # is a deletion, and nothing else is written
    for tracked in (False, True):
        if form == "block":
            package = create_package()
            package.body.insert_xml(
                '<w:sdt><w:sdtPr><w:rPr><w:b/></w:rPr><w:id w:val="4"/></w:sdtPr><w:sdtContent>'
                "<w:p><w:r><w:t>ONE</w:t></w:r></w:p><w:p><w:r><w:t>TWO</w:t></w:r></w:p></w:sdtContent></w:sdt>",
                location="End",
            )
            control = package.body.content_controls[0]
        else:
            package, control = _run_control("<w:r><w:t>VAL</w:t></w:r>", run_properties="<w:rPr><w:b/></w:rPr>")
        if tracked:
            package.change_tracking_mode = "TrackAll"
        control.get_range("Content").delete()
        xml = to_xml(control.element)
        if tracked:
            assert not control.is_showing_placeholder and "<w:delText>" in xml and "     " not in xml, xml
            continue
        assert control.is_showing_placeholder and control.text == "", xml
        assert _filler(control) == [("     ", to_xml(control.run_properties))], xml
        if form == "block":
            assert len(control.paragraphs) == 1
    # in Word's form the emptied control shows its entry again (core-ts check 66 case 3): Phase B's show_placeholder,
    # so here it is left empty, the flag off, until then
    _package, control = _run_control(
        "<w:r><w:t>VAL</w:t></w:r>",
        run_properties=f'<w:placeholder><w:docPart w:val="{DEFAULT_ENTRY}"/></w:placeholder>',
    )
    control.get_range("Content").delete()
    assert (control.text, control.is_showing_placeholder, _filler(control)) == ("", False, [])


def test_a_run_level_controls_whole_range_deleted_takes_it_as_its_replace_does():
    # D3, unmeasured for the delete (CR-019 section 4.10): untracked the control goes; tracked it is marked deleted,
    # nothing inside marked deleted and its flag cleared (core-ts's D3 at d84f3da; check 63's run 2)
    package, control = _run_control(_PLACEHOLDER_RUN, showing=True)
    control.get_range("Whole").delete()
    assert (package.body.paragraphs[0].text, package.body.content_controls) == ("A  B", [])
    package, control = _run_control(_PLACEHOLDER_RUN, showing=True)
    package.change_tracking_mode = "TrackAll"
    control.get_range("Whole").delete()
    xml = to_xml(package.body.paragraphs[0].element)
    assert re.search(
        r"<w:customXmlDelRangeStart [^>]*/><w:sdt><w:sdtPr><w:id w:val=\"2\"/></w:sdtPr><w:sdtContent>"
        r"<w:customXmlDelRangeEnd [^>]*/><w:customXmlDelRangeStart [^>]*/></w:sdtContent></w:sdt>",
        xml,
    ), xml


def _run_control(content: str, *, showing: bool = False, run_properties: str = "") -> tuple[Any, Any]:
    package = create_package()
    flag = "<w:showingPlcHdr/>" if showing else ""
    package.body.insert_xml(
        f'<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:sdt><w:sdtPr>{run_properties}<w:id w:val="2"/>{flag}</w:sdtPr>'
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
# core-ts's tests that are no replay (placeholders.test.mjs at d84f3da): D3, D6 and D10
# ---------------------------------------------------------------------------

#: A run-level control taken whole under tracking: between Word's markers, nothing inside marked deleted, its
#: placeholder gone unmarked and its flag cleared (check 63's run 2; core-ts's D3 at d84f3da).
_MARKED = (
    '<w:customXmlDelRangeStart/><w:sdt><w:sdtPr><w:alias w:val="Check 63 control {n}"/><w:tag w:val="c63-{n}"/>'
    '<w:id w:val="{id}"/></w:sdtPr><w:sdtContent><w:customXmlDelRangeEnd/><w:customXmlDelRangeStart/></w:sdtContent>'
    "</w:sdt><w:customXmlDelRangeEnd/>"
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
    # tracked (44 to 46): Word's markers around the control and inside it, nothing inside marked deleted and the flag
    # cleared, a Replace's insertion first
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
    assert regions[38] == (
        '<w:customXmlDelRangeStart/><w:sdt><w:sdtPr><w:alias w:val="Check 63 control 38"/><w:tag w:val="c63-38"/>'
        '<w:id w:val="63038"/></w:sdtPr><w:sdtContent><w:customXmlDelRangeEnd/><w:p/><w:customXmlDelRangeStart/>'
        "</w:sdtContent></w:sdt><w:customXmlDelRangeEnd/>"
    )
    # a tracked write into the control: only the insertion, plain, the placeholder gone unmarked (D2 (b), D5), in a
    # bound plain-text control too (28, 47)
    assert regions[28] == (
        '<w:sdt><w:sdtPr><w:alias w:val="Check 63 control 28"/><w:tag w:val="c63-28"/><w:id w:val="63028"/></w:sdtPr>'
        "<w:sdtContent><w:p><w:ins><w:r><w:t>NEW28</w:t></w:r></w:ins></w:p></w:sdtContent></w:sdt>"
    )
    assert "<w:sdtContent><w:p><w:ins><w:r><w:t>NEW47</w:t></w:r></w:ins></w:p></w:sdtContent>" in regions[47]
    # a control's Content deleted in the bare form: the flag on over Word's filler, five spaces, untracked or not (D10)
    for n in (6, 32):
        assert (
            "<w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>     </w:t></w:r></w:p></w:sdtContent>"
            in regions[n]
        ), n
    for n in (17, 43):
        assert "<w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:r><w:t>     </w:t></w:r></w:sdtContent>" in regions[n], n
    # accepted, the controls taken go and the text stays; rejected, they come back empty, their flag cleared as Word's
    # markup has it (unmeasured)
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
        False,
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


# ---------------------------------------------------------------------------
# check 67 (core-ts CR-002 sections 63.13 and 63.14, 2bafbcb): check 58's three bare-form writes, in three files, four
# phases and through two lookups, every one a plain run
# ---------------------------------------------------------------------------

#: The three files, their letters and what else each holds (core-ts's ``CHECK67_FILES``).
FILES_67 = [
    ("67a-plain", "a", "nothing else"),
    ("67b-custom-xml", "b", "a custom XML part and a control bound to it"),
    ("67c-revisions", "c", "tracked changes by another author"),
]
PHASES_67 = ["simple-untouched", "all-untouched", "all-off", "simple-off"]


def _cases_67() -> list[dict[str, Any]]:
    """Check 67's cases, the same in every file, as ``build.mjs`` at core-ts ``2bafbcb`` builds them (``check67Cases``)."""
    writes = [("control-text-replace", "block"), ("control-text-end", "block"), ("control-text-end", "run")]
    out: list[dict[str, Any]] = []
    for phase in PHASES_67:
        for lookup in ("tag", "list"):
            for op, shape in writes:
                out.append({"n": len(out) + 1, "op": op, "shape": shape, "phase": phase, "lookup": lookup})
    return out


CASES_67 = _cases_67()


def _view_67(case: dict[str, Any]) -> dict[str, Any]:
    """What both read (core-ts's check 67 ``view``): the digest, the control's type and text, what the write returned."""
    after = case.get("after") if isinstance(case.get("after"), dict) else {}
    control = after.get("control")
    if isinstance(control, dict):
        control = {key: control.get(key) for key in ("type", "text")}
    return {"digest": digest(case), "control": control, "returned": (case.get("result") or {}).get("returned")}


def _report_digest_67(case: dict[str, Any]) -> str:
    return f"{case['n']} CPS - {f'PLAIN{case["n"]}A  PLAIN{case["n"]}B' if case['shape'] == 'run' else ''}"


def _written_digest_67(case: dict[str, Any]) -> str:
    n = case["n"]
    return f"{n} C-- in {f'PLAIN{n}A FILL{n} PLAIN{n}B' if case['shape'] == 'run' else f'FILL{n}'}"


def _controls_67(document: bytes) -> dict[int, str]:
    """Each of a saved main document part's controls tagged ``c67-<n>``, by n, as :func:`_canon` writes it."""
    root = etree.fromstring(document)
    out: dict[int, str] = {}
    for sdt in root.iter(f"{{{W}}}sdt"):
        tag = sdt.find(f"{{{W}}}sdtPr/{{{W}}}tag")
        found = re.match(r"^c67-(\d+)$", tag.get(f"{{{W}}}val", "") if tag is not None else "")
        if found and found.group(1) != "0":
            out[int(found.group(1))] = _canon(sdt)
    return out


def test_check_67s_files_and_snippet_agree_each_control_in_the_bare_form_showing_its_placeholder_the_files_differ_only_in_what_else_they_hold():
    for stem, letter, holds in FILES_67:
        package = load(OFFICE_JS / f"{stem}.docx")
        body = package.body
        assert [r for p in body.paragraphs if (r := checks.parse_region_67(p.text)) is not None] == CASES_67, stem
        assert body.paragraphs[0].text.startswith(f"Check 67{letter}: ") and holds in body.paragraphs[0].text
        controls = [c for c in body.get_content_controls() if c.tag != "c67-0"]
        assert [c.tag for c in controls] == [f"c67-{c['n']}" for c in CASES_67], stem
        assert all(c.is_showing_placeholder and c.placeholder_entry is None and c.text == "" for c in controls), stem
        names = zipfile.ZipFile(OFFICE_JS / f"{stem}.docx").namelist()
        # what else each holds: 67b a custom XML part and a bound control, 67c tracked changes in its first paragraph
        assert any(name.startswith("customXml/") for name in names) == (letter == "b"), stem
        assert any(c.tag == "c67-0" for c in body.get_content_controls()) == (letter == "b"), stem
        assert [(c.type, c.text) for c in body.get_tracked_changes()] == (
            [("Added", "Inserted."), ("Deleted", " Deleted.")] if letter == "c" else []
        ), stem
        # the file is the one Word read: Word's save of it as opened reads paragraph for paragraph and control for
        # control as it does
        opened = load(OFFICE_JS / f"67{letter}-opened-word365.docx").body
        assert [p.text for p in opened.paragraphs] == [p.text for p in body.paragraphs], stem
        assert [(c.tag, c.is_showing_placeholder, c.text) for c in opened.get_content_controls()] == [
            (c.tag, c.is_showing_placeholder, c.text) for c in body.get_content_controls()
        ], stem
    assert [name for name, _view in checks.PHASES_67] == PHASES_67


def test_check_67s_run_words_jsons_are_of_these_files_cases_every_step_run_in_order_and_its_saves_are_words_as_copied():
    manifest = dict(
        reversed(line.split("  ", 1))
        for line in (OFFICE_JS / "SHA1SUMS").read_text(encoding="utf-8").splitlines()
        if line
    )
    for name, sha1 in SHA1_67.items():
        assert manifest[name] == sha1, name
        assert hashlib.sha1((OFFICE_JS / name).read_bytes()).hexdigest() == sha1, name
    for stem, letter, _holds in FILES_67:
        measured = word(f"{stem}-word365.json")
        assert [measured["check"], measured["host"]["version"], list(measured["steps"])] == [
            67,
            "16.0.20326.20232",
            ["report", *PHASES_67],
        ]
        assert measured["title"].startswith(f"Check 67{letter}: ")
        report = measured["steps"]["report"]["cases"]
        assert [{k: c[k] for k in ("n", "op", "shape", "phase", "lookup")} for c in report] == CASES_67
        for phase in PHASES_67:
            assert [c["n"] for c in measured["steps"][phase]["cases"]] == [
                c["n"] for c in CASES_67 if c["phase"] == phase
            ]


def test_word_365s_answers_check_67_every_write_into_a_control_in_the_bare_form_is_a_plain_run_the_flag_cleared():
    """core-ts: "Word 365's answers (check 67): every write into a control in the bare form is a plain run, the flag
    cleared, in every phase, lookup and file, as check 58's were" (its section 63.14)."""
    for stem, letter, _holds in FILES_67:
        measured = word(f"{stem}-word365.json")
        report = measured["steps"]["report"]["cases"]
        assert [digest(c) for c in report] == [_report_digest_67(c) for c in CASES_67], stem
        for phase in PHASES_67:
            cases = measured["steps"][phase]["cases"]
            assert not [c for c in cases if c.get("error")], f"{stem} {phase}"
            # no Placeholder Text style, no w:color auto: the text plain, as in Word's own form (check 64)
            assert [digest(c) for c in cases] == [_written_digest_67(c) for c in cases], f"{stem} {phase}"
            for c in cases:
                assert not re.search(r'w:color w:val="auto"|PlaceholderText', c["after"]["ooxml"]), f"{stem} {c['n']}"
        # Word's save holds the same: one plain run in each control, the flag gone
        saved = _controls_67(zipfile.ZipFile(OFFICE_JS / f"67{letter}-run-word365.docx").read("word/document.xml"))
        for c in CASES_67:
            plain = f"<w:r><w:t>FILL{c['n']}</w:t></w:r>"
            content = f"<w:p>{plain}</w:p>" if c["shape"] == "block" else plain
            assert saved[c["n"]].endswith(f"</w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>"), (
                f"{stem} {c['n']}"
            )
            assert "showingPlcHdr" not in saved[c["n"]], f"{stem} {c['n']}"
    # only check 63's run 1 saves carry w:formsDesign among the placeholder checks' here (check 63's, 64's and 67's)
    designed = sorted(
        path.name
        for path in OFFICE_JS.glob("6[347]*.docx")
        if "<w:formsDesign/>" in zipfile.ZipFile(path).read("word/settings.xml").decode("utf-8")
    )
    assert designed == ["63-placeholders-word365.docx", "63-run-word365.docx"]


def test_check_67s_run_through_the_port_is_word_365s_in_every_step_all_72_writes_a_plain_run():
    """core-ts: "the shim agrees with Word in all 72 of check 67's writes" (its section 63.17)."""
    written = 0
    for stem, _letter, _holds in FILES_67:
        package = load(OFFICE_JS / f"{stem}.docx")
        measured = word(f"{stem}-word365.json")
        differing: dict[str, list[int]] = {}
        for step in ["report", *PHASES_67]:
            ours = checks.run_check_67(package, step)
            theirs = measured["steps"][step]
            assert [ours["check"], ours["step"], ours["caseCount"]] == [67, step, theirs["caseCount"]]
            # the title as Word read it, but in All Markup, where Word's paragraph text holds 67c's deleted word too: a
            # view this package's text doesn't take (its default is Simple Markup's)
            title = theirs["title"]
            if step.startswith("all-") and stem == "67c-revisions":
                assert title.endswith(" Inserted. Deleted.")
                title = title.removesuffix(" Deleted.")
            assert ours["title"] == title
            assert [c["n"] for c in ours["cases"]] == [c["n"] for c in theirs["cases"]]
            differing[step] = [
                c["n"] for c, mine in zip(theirs["cases"], ours["cases"], strict=True) if _view_67(c) != _view_67(mine)
            ]
            if step != "report":
                written += len(ours["cases"])
                assert [digest(c) for c in ours["cases"]] == [_written_digest_67(c) for c in ours["cases"]]
        assert differing == {step: [] for step in ["report", *PHASES_67]}, stem
        assert package.change_tracking_mode == "Off"
    assert written == 72


def test_check_67s_saved_markup_is_words_save_after_the_run_control_for_control():
    for stem, letter, _holds in FILES_67:
        package = load(OFFICE_JS / f"{stem}.docx")
        for step in ["report", *PHASES_67]:
            checks.run_check_67(package, step)
        ours = _controls_67(zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml"))
        theirs = _controls_67(zipfile.ZipFile(OFFICE_JS / f"67{letter}-run-word365.docx").read("word/document.xml"))
        assert sorted(ours) == sorted(theirs) == [c["n"] for c in CASES_67], stem
        assert [n for n in theirs if ours[n] != theirs[n]] == [], stem
