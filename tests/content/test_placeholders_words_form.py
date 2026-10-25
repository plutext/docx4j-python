"""Placeholders in Word's form: CR-019 Phase B, the port of core-ts's CR-002 sections 63.7 to 63.12 and 63.18.

    .venv-fork/bin/python -m pytest tests/content/test_placeholders_words_form.py

docx4j-core-ts ran Word 365 checks 65 and 66 on what Word needs of a
placeholder in its own form, a ``w:placeholder`` naming a glossary entry (its
sections 63.7 to 63.11; Word 365 16.0.20326.20232, 2026-10-25), and
implemented the answers, D8 and D9, at ``020c039``, revised at ``aefc41f`` (its
section 63.18: a placeholder shows a copy of its entry's content; a text
control naming a missing entry is given Word's default). Its
``test/placeholders.test.mjs`` replays them through its Office JS shim; here
``tests/office_js/checks.py`` ports the two snippets step for step
(:func:`run_check_65`, :func:`run_check_66`), and this file holds this
package's run to Word's, case by case, through core-ts's ``digest`` and
``view``, and ``placeholderText`` too, which Phase B reads from the glossary.

* **Check 65:** ``65a-reference-only`` names the default entry with no
  glossary part; ``65b-minimal-glossary`` has a glossary of the document part
  alone, an entry of its own, and a control naming an entry it lacks.
* **Check 66:** ``66a-glossary-with-styles`` has the glossary document and
  styles parts this package writes (D9 (a)); ``66b-word-glossary`` Word's six.

What still differs is pinned, with its reason: 65b's case 9 (Word's glossary
had lost the style), and in Word's saves the references Word's own session
dropped and the controls typed into by hand. Then core-ts's tests that are no
replay: D8's setter, binding and made controls; D9's glossary made, added to,
and left untouched; section 63.18's entry content and missing entry.
"""

from __future__ import annotations

import hashlib
import importlib.util
import io
import re
import zipfile
from typing import Any

import pytest
from conftest import ROOT

from docx4j_py import create_package, load
from docx4j_py.model.content.placeholders import (
    DEFAULT_PLACEHOLDER_ENTRY,
    DEFAULT_PLACEHOLDER_TEXT,
    GlossaryEntries,
    entry_run_properties,
)
from docx4j_py.traversal import text_of
from docx4j_py.wml import to_xml

OFFICE_JS = ROOT / "tests" / "office_js"

_spec = importlib.util.spec_from_file_location("office_js_checks", OFFICE_JS / "checks.py")
checks = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(checks)

_writes_spec = importlib.util.spec_from_file_location("placeholder_writes", ROOT / "tests" / "content" / "test_placeholder_writes.py")
writes = importlib.util.module_from_spec(_writes_spec)
assert _writes_spec.loader is not None
_writes_spec.loader.exec_module(writes)
digest, view, word = writes.digest, writes.view, writes.word

PLACEHOLDER = DEFAULT_PLACEHOLDER_TEXT
DEFAULT_ENTRY = DEFAULT_PLACEHOLDER_ENTRY
#: 65b's own entry, named as Word names one, and its text; an entry no glossary holds (check65/build.mjs).
CUSTOM_ENTRY = "65656565656545658565656565650001"
CUSTOM_TEXT = "Enter the name."
MISSING_ENTRY = "65656565656545658565656565650002"

#: The files, as core-ts aefc41f has them: each check's two files (git-ignored there, written by build.mjs before
#: Word's run: 2026-10-24 22:20 for check 65's run at 05:40, 2026-10-25 06:07 for check 66's at 06:28), Word's
#: JSONs, and Word's saves, which test/check65/ and test/check66/'s manifest.json list with these SHA-1s.
SHA1 = {
    "65a-reference-only.docx": "9ce50cac4db29c6863a0da45e731e3ed74eff812",
    "65b-minimal-glossary.docx": "803000981cf7aa3d98a98ad07a3da14f1809594a",
    "65a-reference-only-word365-report.json": "7af5f809a2ccd9fec787fbb5e974029727cb5411",
    "65a-reference-only-word365-run.json": "a6b4d64e5b303b210fe0d1e0ce0f43a8b59126c9",
    "65b-minimal-glossary-word365-report.json": "e1bd6d75b57c916c6bb342e06c6650cc71bfd498",
    "65b-minimal-glossary-word365-run.json": "86cc450d2e71e44b1601ed190eb8f857334ce098",
    "65a-opened-word365.docx": "62099bc3c8f03c831db1c30f89c5b1a916d628cc",
    "65a-run-word365.docx": "103732a66b418219820020d50fcd1f8b77fbc02c",
    "65b-opened-word365.docx": "943147384f4f97483887e903a7eaf1b26258e198",
    "65b-run-word365.docx": "6d83e00bb110655256035652f05c640c36434472",
    "65b-reopened-typed-word365.docx": "483ccf6e6851a8d82124994570724d0108e3bd34",
    "66a-glossary-with-styles.docx": "2f1b3ffd98d6a947a2c0b5e5004fed662e97b91c",
    "66b-word-glossary.docx": "daa6b8ce9c0ef5ded1744155e155ce20ce0db356",
    "66a-glossary-with-styles-word365-report.json": "f704ae9a4dac82a6cb6bb5206d1109a165ca8173",
    "66a-glossary-with-styles-word365-run.json": "b5d755cdb896edde6d1ed516f2ad32a619134343",
    "66b-word-glossary-word365-report.json": "57bd5370c454f0538829392b6bdc513e7ab6e7ae",
    "66b-word-glossary-word365-run.json": "2c0cc39a0ef0625b80d2f29928aff224926c575a",
    "66a-opened-word365.docx": "edfd58eae3231210d1a33e5a2f4893807600d36c",
    "66a-run-word365.docx": "4c47d2aa764314ac134ebb68ea3154555f1add0b",
    "66a-reopened-word365.docx": "b4f371fd0ef5f3b4dcf94d7077f39896801dea62",
    "66a-typed-word365.docx": "59c34e08dd25edb91812c9ce233513d653a1a9e7",
    "66b-opened-word365.docx": "343f9e002eaa3a0454a77e56ca0aef8777dc5cf9",
    "66b-run-word365.docx": "372254cc3333419a4114529c9a35a579f03d0c38",
    "66b-reopened-word365.docx": "3e4e23751a24f0eec62e9e18faf5b6062a385cbc",
    "66b-typed-word365.docx": "d9309b2e7a002f5e105b72cc621d1bdfdfa4882b",
}

FILES_65 = ["65a-reference-only", "65b-minimal-glossary"]
FILES_66 = ["66a-glossary-with-styles", "66b-word-glossary"]


def _cases_65(name: str) -> list[dict[str, Any]]:
    """A check 65 file's cases as ``build.mjs`` at core-ts ``aefc41f`` builds them (``check65Cases``)."""
    out: list[dict[str, Any]] = []

    def add(op: str, shape: str, mode: str = "untracked") -> None:
        out.append({"n": len(out) + 1, "op": op, "shape": shape, "mode": mode})

    add("control-text-end", "block")
    add("control-paragraph-end", "block")
    add("control-text-replace", "block", "tracked")
    add("control-text-replace", "block-bound")
    add("content-text-replace", "run-bound")
    add("set-placeholder-text", "block")
    add("open", "block")
    add("open", "block-bound")
    add("make", "made-block")
    if name.startswith("65b"):
        add("control-text-end", "custom")
        add("open", "custom")
        add("control-text-end", "missing")
        add("open", "missing")
    add("type", "bare-bound")
    return out


#: Check 66's six cases, the same in both files (check66/build.mjs's ``check66Cases``).
CASES_66 = [
    {"n": n, "op": op, "shape": shape, "mode": "untracked"}
    for n, (op, shape) in enumerate(
        [
            ("open", "block"),
            ("type", "block-bound"),
            ("content-delete", "block-value"),
            ("make", "made-block"),
            ("set-placeholder-text", "block"),
            ("type", "bare-bound"),
        ],
        start=1,
    )
]


def run(check: int, name: str, mode: str = "run") -> tuple[Any, dict[str, Any]]:
    package = load(OFFICE_JS / f"{name}.docx")
    return package, getattr(checks, f"run_check_{check}")(package, mode)


def _parts(data: bytes) -> dict[str, str]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name).decode("utf-8") for name in archive.namelist()}


def _save(name: str) -> dict[str, str]:
    return _parts((OFFICE_JS / name).read_bytes())


def _controls(xml: str, check: int) -> dict[int, str]:
    """Every control in a saved main document part: ``<n> <flag><style> <entry> <text>`` (core-ts's ``controls66``).

    The entry is ``default``, ``own`` (32 hex digits), ``-`` or the name.
    """
    out: dict[int, str] = {}
    for sdt in re.findall(r"<w:sdt>[\s\S]*?</w:sdt>", xml):
        found = re.search(r'<w:docPart w:val="([^"]+)"/>', sdt)
        entry = found.group(1) if found else None
        content = re.search(r"<w:sdtContent>[\s\S]*</w:sdtContent>", sdt)
        text = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", content.group(0) if content else ""))
        shown = (
            "-"
            if entry is None
            else "default"
            if entry == DEFAULT_ENTRY
            else "own"
            if re.fullmatch(r"[0-9A-F]{32}", entry)
            else entry
        )
        flag = "P" if "<w:showingPlcHdr/>" in sdt else "-"
        style = "S" if '<w:rStyle w:val="PlaceholderText"/>' in sdt else "-"
        n = int(re.search(rf"c{check}-(\d+)", sdt).group(1))
        out[n] = f"{n} {flag}{style} {shown} {text}"
    return out


# ---------------------------------------------------------------------------
# the files, and the port's agreement with the snippets
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
    for check, names, cases in ((65, FILES_65, _cases_65), (66, FILES_66, lambda _name: CASES_66)):
        for name in names:
            for mode in ("report", "run"):
                measured = word(f"{name}-word365-{mode}.json")
                expected = cases(name)
                assert [measured["check"], measured["mode"], measured["host"]["version"], measured["caseCount"]] == [
                    check,
                    mode,
                    "16.0.20326.20232",
                    len(expected),
                ]
                assert [{k: c[k] for k in ("n", "op", "shape", "mode")} for c in measured["cases"]] == expected


def test_check_65s_files_and_snippet_agree_65a_has_no_glossary_part_65b_one_of_one_part():
    assert [len(_cases_65(name)) for name in FILES_65] == [10, 14]
    for name in FILES_65:
        package = load(OFFICE_JS / f"{name}.docx")
        body = package.body
        cases = _cases_65(name)
        assert [r for p in body.paragraphs if (r := checks.parse_region_65(p.text)) is not None] == cases
        assert body.paragraphs[0].text.startswith(f"Check 65{name[2]}: ")
        with_control = [c for c in cases if c["shape"] != "made-block"]
        assert [c.tag for c in body.content_controls] == [f"c65-{c['n']}" for c in with_control]
        assert all(c.is_showing_placeholder for c in body.content_controls)
        # every control names its entry but the bare-form one typed into by hand
        named = {"custom": CUSTOM_ENTRY, "missing": MISSING_ENTRY}
        assert [c.placeholder_entry for c in body.content_controls] == [
            None if c["shape"] == "bare-bound" else named.get(c["shape"], DEFAULT_ENTRY) for c in with_control
        ]
        assert body.content_controls[-1].xml_mapping.is_mapped
        glossary = [k for k in _parts((OFFICE_JS / f"{name}.docx").read_bytes()) if k.startswith("word/glossary/")]
        if name.startswith("65a"):
            assert glossary == [] and package.main_document_part.glossary_document_part is None
        else:
            assert glossary == ["word/glossary/document.xml"]
            assert package.main_document_part.glossary_document_part is not None


def test_check_66s_files_and_snippet_agree_66a_has_the_document_and_styles_parts_66b_words_six():
    for name in FILES_66:
        package = load(OFFICE_JS / f"{name}.docx")
        body = package.body
        assert [r for p in body.paragraphs if (r := checks.parse_region_65(p.text)) is not None] == CASES_66
        assert body.paragraphs[0].text.startswith(f"Check 66{name[2]}: ")
        with_control = [c for c in CASES_66 if c["shape"] != "made-block"]
        assert [c.tag for c in body.content_controls] == [f"c66-{c['n']}" for c in with_control]
        assert [c.is_showing_placeholder for c in body.content_controls] == [
            c["shape"] != "block-value" for c in with_control
        ]
        assert [c.placeholder_entry for c in body.content_controls] == [DEFAULT_ENTRY] * 4 + [None]
        parts = _parts((OFFICE_JS / f"{name}.docx").read_bytes())
        glossary = sorted(k for k in parts if k.startswith("word/glossary/"))
        if name.startswith("66a"):
            assert glossary == ["word/glossary/_rels/document.xml.rels", "word/glossary/document.xml", "word/glossary/styles.xml"]
            assert re.search(
                r'<w:docPartBody><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>'
                r'<w:t xml:space="preserve">Click or tap here to enter text\.</w:t></w:r></w:p></w:docPartBody>',
                parts["word/glossary/document.xml"],
            )
        else:
            assert len(glossary) == 6


# ---------------------------------------------------------------------------
# Word's runs, held to case by case
# ---------------------------------------------------------------------------

#: Word 365's runs of check 65 (2026-10-25), as ``digest`` reads them: core-ts's ``WORD_RUN_65``, check 64's answers
#: in both files. 65b's made control (9) shows its placeholder without the style, which Word dropped from the
#: glossary it read.
WORD_RUN_65 = {
    "65a-reference-only": [
        "1 C-- in NEW1",
        "2 C-- in |NEW2",
        "3 C-- in NEW3",
        "4 C-- in NEW4",
        "5 C-- in PLAIN5A NEW5 PLAIN5B",
        "6 CPS - ",
        "7 CPS - ",
        "8 CPS - ",
        "9 CPS - ",
        "10 CPS - ",
    ],
    "65b-minimal-glossary": [
        "1 C-- in NEW1",
        "2 C-- in |NEW2",
        "3 C-- in NEW3",
        "4 C-- in NEW4",
        "5 C-- in PLAIN5A NEW5 PLAIN5B",
        "6 CPS - ",
        "7 CPS - ",
        "8 CPS - ",
        "9 CP- - ",
        "10 C-- in NEW10",
        "11 CPS - ",
        "12 C-- in NEW12",
        "13 CPS - ",
        "14 CPS - ",
    ],
}

#: Check 65's one write differing from Word's: 65b's made control, which Word showed unstyled, Word's reading of a
#: glossary whose styles lack Placeholder Text having dropped the style; this package keeps a style a glossary does
#: not define, and its own glossaries define it (D9; core-ts section 63.18).
DIFFERS_65 = {("65b-minimal-glossary", 9): "9 CPS - "}


def test_word_365s_answers_check_65_the_reference_gives_words_form_and_placeholder_text_reads_the_entry():
    for name in FILES_65:
        for case in word(f"{name}-word365-report.json")["cases"]:
            control = case["after"]["control"]
            if not control:
                continue
            missing = name.startswith("65a") or case["shape"] in ("missing", "bare-bound")
            entry = None if missing else CUSTOM_TEXT if case["shape"] == "custom" else PLACEHOLDER
            assert control["placeholderText"] == entry, (name, case["n"])
        measured = word(f"{name}-word365-run.json")
        assert [digest(c) for c in measured["cases"]] == WORD_RUN_65[name]
        assert not [c for c in measured["cases"] if c.get("error")]
        assert "<w:ins " in measured["cases"][2]["after"]["ooxml"]
        assert not re.search(r"<w:del |<w:delText", measured["cases"][2]["after"]["ooxml"])
        assert [measured["cases"][3]["after"]["node"], measured["cases"][4]["after"]["node"]] == ["NEW4", "NEW5"]
        assert measured["cases"][5]["result"]["placeholderText"] == "Enter the name 6."


@pytest.mark.parametrize("name", FILES_65)
@pytest.mark.parametrize("mode", ["report", "run"])
def test_check_65s_run_through_the_port_is_word_365s_but_65bs_made_control(name, mode):
    _package, ours = run(65, name, mode)
    measured = word(f"{name}-word365-{mode}.json")
    assert ours["caseCount"] == len(_cases_65(name))
    agreeing = []
    for theirs, mine in zip(measured["cases"], ours["cases"], strict=True):
        n = theirs["n"]
        expected = view(theirs)
        if mode == "run" and (name, n) in DIFFERS_65:
            expected = {**expected, "digest": DIFFERS_65[name, n]}
            assert view(theirs) != expected
        else:
            agreeing.append(n)
        assert view(mine) == expected, n
        # placeholderText reads the entry the control names, None where the glossary lacks it, as Word's (CR-019 Q2;
        # until Phase B the runs were read); the setter's read back too
        control = theirs["after"].get("control")
        if isinstance(control, dict):
            assert mine["after"]["control"]["placeholderText"] == control["placeholderText"], n
        assert (mine.get("result") or {}).get("placeholderText") == (theirs.get("result") or {}).get("placeholderText")
    # at caff8cf's code: every report read but placeholderText (in 9 and 3 controls), and every write but 9 of each
    # file, placeholderText aside (in 9 and 10 controls)
    assert len(agreeing) == len(measured["cases"]) - (1 if mode == "run" and name.startswith("65b") else 0)


@pytest.mark.parametrize("name", FILES_66)
@pytest.mark.parametrize("mode", ["report", "run"])
def test_check_66s_run_through_the_port_is_word_365s_in_every_case(name, mode):
    _package, ours = run(66, name, mode)
    measured = word(f"{name}-word365-{mode}.json")
    assert ours["caseCount"] == 6
    # at caff8cf's code the emptied control (3) and the made one (4) showed no placeholder, and placeholderText read
    # the runs (3, 4 and 6 differing)
    for theirs, mine in zip(measured["cases"], ours["cases"], strict=True):
        assert view(mine) == view(theirs), theirs["n"]
        control = theirs["after"].get("control")
        if isinstance(control, dict):
            assert mine["after"]["control"]["placeholderText"] == control["placeholderText"], theirs["n"]
        assert (mine.get("result") or {}).get("placeholderText") == (theirs.get("result") or {}).get("placeholderText")
    expected = ["1 CPS - ", "2 CPS - ", "3 C-- - VALUE3", "4 x-- - ", "5 CPS - ", "6 CPS - "]
    if mode == "run":
        expected = ["1 CPS - ", "2 CPS - ", "3 CPS - ", "4 CPS - ", "5 CPS - ", "6 CPS - "]
    assert [digest(c) for c in ours["cases"]] == expected


# ---------------------------------------------------------------------------
# Word's saves after the run, control by control
# ---------------------------------------------------------------------------

DROPPED = "Word's session dropped a reference to an entry no glossary held when it saved the file as opened; this run keeps it"
DROPPED_MISSING = DROPPED + ", writing nothing under the missing name (core-ts section 63.18)"
TYPED = "typed into by hand after the run (TYPED10), which the snippet does not do"
UNSTYLED = "Word showed the entry unstyled, its glossary's styles lacking Placeholder Text (as the run, 65b case 9)"

#: Each file's controls whose markup in this package's save after the run differs from Word's: ``n: why``.
SAVED_DIFFER = {
    "65a-reference-only": {**{n: DROPPED + ", writing the default entry (D9)" for n in (1, 2, 3, 4, 5, 7, 8)}, 10: TYPED},
    "65b-minimal-glossary": {9: UNSTYLED, 10: TYPED, 12: DROPPED_MISSING, 13: DROPPED_MISSING},
    "66a-glossary-with-styles": {},
    "66b-word-glossary": {},
}


@pytest.mark.parametrize("name", [*FILES_65, *FILES_66])
def test_the_saved_controls_are_words_save_after_the_run_but_for_its_sessions_own(name):
    check = int(name[:2])
    package, _ours = run(check, name)
    ours = _controls(_parts(package.save())["word/document.xml"], check)
    theirs = _controls(_save(f"{name[:3]}-run-word365.docx")["word/document.xml"], check)
    assert sorted(ours) == sorted(theirs)
    differing = {n for n in theirs if ours[n] != theirs[n]}
    assert differing == set(SAVED_DIFFER[name]), {n: (theirs[n], ours[n]) for n in differing}
    # 66's every control, Word's: the emptied one and the made one showing the default entry, grey, and 5 its own
    if check == 66:
        assert [ours[n] for n in (3, 4, 5)] == [
            f"3 PS default {PLACEHOLDER}",
            f"4 PS default {PLACEHOLDER}",
            "5 PS own Enter the name 5.",
        ]


def test_word_365s_saves_check_65_only_an_entry_keeps_its_reference_and_a_glossary_without_its_styles_loses_the_style():
    opened_a = _save("65a-opened-word365.docx")
    assert "<w:placeholder>" not in opened_a["word/document.xml"]
    assert not [k for k in opened_a if k.startswith("word/glossary/")]
    opened_b = _save("65b-opened-word365.docx")
    named = _controls(opened_b["word/document.xml"], 65)
    assert [n for n, line in named.items() if line.split()[2] != "-"] == [1, 2, 3, 4, 5, 6, 7, 8, 10, 11]
    assert sorted(k for k in opened_b if k.startswith("word/glossary/")) == [
        "word/glossary/_rels/document.xml.rels",
        "word/glossary/document.xml",
        "word/glossary/fontTable.xml",
        "word/glossary/settings.xml",
        "word/glossary/styles.xml",
        "word/glossary/webSettings.xml",
    ]
    assert 'w:styleId="PlaceholderText"' not in opened_b["word/glossary/styles.xml"]
    # text of a control's own (case 6) is a new entry: 32 hex digits, the default's docPartPr, its text in the style
    for save in (_save("65a-run-word365.docx"), _save("65b-run-word365.docx")):
        name = re.search(
            r'<w:tag w:val="c65-6"/><w:id w:val="65006"/><w:placeholder><w:docPart w:val="([0-9A-F]{32})"/>',
            save["word/document.xml"],
        ).group(1)
        assert re.search(
            rf'<w:docPartPr><w:name w:val="{name}"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/>'
            r'</w:category><w:types><w:type w:val="bbPlcHdr"/></w:types><w:behaviors><w:behavior w:val="content"/>'
            r'</w:behaviors><w:guid w:val="\{[0-9A-F-]{36}\}"/></w:docPartPr>',
            save["word/glossary/document.xml"],
        )
    typed = _save("65b-reopened-typed-word365.docx")
    assert "<value14>TYPED14</value14>" in typed["customXml/item1.xml"]


def test_word_365s_saves_check_66_a_glossary_of_the_document_and_styles_parts_does_what_words_six_parts_do():
    strip = re.compile(r' w:rsid\w*="[^"]*"|<w:rsid [^>]*/>|<w:rsids>[\s\S]*?</w:rsids>')
    for letter in "ab":
        saved = {step: _save(f"66{letter}-{step}-word365.docx") for step in ("opened", "run", "reopened", "typed")}
        after_run = [
            f"1 PS default {PLACEHOLDER}",
            f"2 PS default {PLACEHOLDER}",
            f"3 PS default {PLACEHOLDER}",
            f"4 PS default {PLACEHOLDER}",
            "5 PS own Enter the name 5.",
            f"6 PS - {PLACEHOLDER}",
        ]
        assert list(_controls(saved["run"]["word/document.xml"], 66).values()) == after_run
        assert list(_controls(saved["reopened"]["word/document.xml"], 66).values()) == after_run
        assert list(_controls(saved["typed"]["word/document.xml"], 66).values()) == [
            after_run[0],
            "2 -- default TYPED2",
            *after_run[2:5],
            "6 -- - TYPED6",
        ]
        assert re.search(r"<value2>TYPED2</value2><value6>TYPED6</value6>", saved["typed"]["customXml/item1.xml"])
        for save in saved.values():
            assert re.search(
                r'<w:docPartBody><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>'
                r"<w:t>Click or tap here to enter text\.</w:t></w:r></w:p></w:docPartBody>",
                strip.sub("", save["word/glossary/document.xml"]),
            )
            assert re.findall(r'w:styleId="([^"]+)"', save["word/glossary/styles.xml"]) == [
                "Normal",
                "DefaultParagraphFont",
                "TableNormal",
                "NoList",
                "PlaceholderText",
            ]


# ---------------------------------------------------------------------------
# core-ts's tests that are no replay: D8 and D9 (its section 63.12), and section 63.18
# ---------------------------------------------------------------------------


def _check_63() -> Any:
    return load(OFFICE_JS / "63-placeholders.docx")


def _tagged(package: Any, *tags: str) -> list[Any]:
    controls = {c.tag: c for c in package.body.content_controls}
    return [controls[tag] for tag in tags]


def test_d8_the_setter_names_an_entry_an_empty_bound_value_shows_the_default_entry_a_made_control_shows_it():
    package = _check_63()
    block, run_control = _tagged(package, "c63-1", "c63-13")
    # a control in the bare form names no entry, and reads None, as Word's (check 63)
    assert [block.placeholder_entry, block.placeholder_text] == [None, None]
    # Word's default text names the default entry; any other a new entry of its own, 32 hex digits (check 65)
    block.placeholder_text = PLACEHOLDER
    run_control.placeholder_text = "Enter a name."
    assert block.placeholder_entry == DEFAULT_ENTRY
    assert re.fullmatch(r"[0-9A-F]{32}", run_control.placeholder_entry)
    assert [block.placeholder_text, run_control.placeholder_text, run_control.is_showing_placeholder, run_control.text] == [
        PLACEHOLDER,
        "Enter a name.",
        True,
        "",
    ]
    xml = run_control.get_xml()
    assert re.search(r'<w:placeholder><w:docPart w:val="[0-9A-F]{32}"/></w:placeholder><w:showingPlcHdr/>', xml)
    assert re.search(
        r'<w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Enter a name\.</w:t></w:r></w:sdtContent>',
        xml,
    )
    # set again, a new entry; the one named before stays
    first = run_control.placeholder_entry
    run_control.placeholder_text = "Enter a name."
    assert run_control.placeholder_entry != first
    assert package.glossary_entries.text_of(first) == "Enter a name."

    # an empty bound value: the bound controls show the default entry, in Word's form, where docx4j's binding writes
    # the bare form (check 64's controls, as Word saves them)
    bound = _check_63()
    bound.custom_xml_parts.apply_bindings()
    controls = [c for c in bound.body.content_controls if c.xml_mapping.is_mapped]
    assert len(controls) == 12
    for control in controls:
        assert [control.placeholder_entry, control.placeholder_text, control.is_showing_placeholder, control.text] == [
            DEFAULT_ENTRY,
            PLACEHOLDER,
            True,
            "",
        ], control.tag
        assert re.search(
            r'<w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click or tap here to enter text\.</w:t>', control.get_xml()
        ), control.tag

    # a control made on an empty range shows the default entry; a kind Word gives another entry is made as before
    made = create_package()
    body = made.body
    point = body.insert_paragraph("Name: ").get_range("End").insert_content_control()
    around = body.insert_paragraph("").insert_content_control("PlainText")
    checkbox = body.insert_paragraph("").insert_content_control("CheckBox")
    assert [point.placeholder_text, point.is_showing_placeholder, around.placeholder_text, around.is_showing_placeholder] == [
        PLACEHOLDER,
        True,
        PLACEHOLDER,
        True,
    ]
    assert [checkbox.placeholder_entry, checkbox.is_showing_placeholder] == [None, False]
    # a control around a paragraph holding text holds that text; one made at a point stands there, the run split
    filled = body.insert_paragraph("Kept.").insert_content_control()
    assert [filled.is_showing_placeholder, filled.text] == [False, "Kept."]
    assert body.paragraphs[0].text == "Name: " and point.parent_paragraph().element is body.paragraphs[0].element
    middle = body.insert_paragraph("AB")
    split = middle.search("A")[0].get_range("After").insert_content_control()
    # until CR-019 Phase B the control wrapped the run around the point (section 10.9 item 1)
    assert (middle.text, split.text, split.is_showing_placeholder) == ("AB", "", True)
    xml = to_xml(middle.element)
    assert xml.index(">A</w:t>") < xml.index("<w:sdt>") < xml.index(">B</w:t>")


def test_d9_a_placeholder_written_where_the_file_has_no_glossary_makes_one_on_save():
    # check 63's file has no glossary
    package = _check_63()
    block, run_control = _tagged(package, "c63-1", "c63-13")
    block.placeholder_text = PLACEHOLDER
    run_control.placeholder_text = "Enter a name."
    saved = package.save()
    parts = _parts(saved)
    assert sorted(k for k in parts if k.startswith("word/glossary/")) == [
        "word/glossary/_rels/document.xml.rels",
        "word/glossary/document.xml",
        "word/glossary/styles.xml",
    ]
    assert package.last_save.added == ["/word/glossary/document.xml", "/word/glossary/styles.xml"]
    assert re.search(
        r'Target="glossary/document\.xml" Type="http://schemas\.openxmlformats\.org/officeDocument/2006/relationships/glossaryDocument"',
        parts["word/_rels/document.xml.rels"],
    )
    assert re.search(
        r'Target="styles\.xml" Type="http://schemas\.openxmlformats\.org/officeDocument/2006/relationships/styles"',
        parts["word/glossary/_rels/document.xml.rels"],
    )
    types = parts["[Content_Types].xml"]
    assert (
        '<Override PartName="/word/glossary/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>' in types
    )
    assert (
        '<Override PartName="/word/glossary/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' in types
    )

    # each entry as Word writes one (check 65): the default entry's docPartPr, a guid of its own, one run in the style
    def entry(name: str, text: str) -> str:
        return (
            rf'<w:docPart><w:docPartPr><w:name w:val="{name}"/><w:category><w:name w:val="General"/>'
            r'<w:gallery w:val="placeholder"/></w:category><w:types><w:type w:val="bbPlcHdr"/></w:types>'
            r'<w:behaviors><w:behavior w:val="content"/></w:behaviors><w:guid w:val="\{[0-9A-F-]{36}\}"/></w:docPartPr>'
            rf'<w:docPartBody><w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>{re.escape(text)}</w:t>'
            r"</w:r></w:p></w:docPartBody></w:docPart>"
        )

    glossary = parts["word/glossary/document.xml"]
    assert re.search(entry(DEFAULT_ENTRY, PLACEHOLDER), glossary)
    assert re.search(entry(run_control.placeholder_entry, "Enter a name."), glossary)
    guids = re.findall(r'<w:guid w:val="([^"]+)"/>', glossary)
    assert len(set(guids)) == 2
    # the styles as Word defines them (check 66's: Word kept such a glossary's entries in their style)
    styles = parts["word/glossary/styles.xml"]
    assert (
        '<w:style w:type="character" w:styleId="DefaultParagraphFont" w:default="1"><w:name w:val="Default Paragraph Font"/>'
        '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' in styles
    )
    assert (
        '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/>'
        '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'
        '<w:rPr><w:color w:val="666666"/></w:rPr></w:style>' in styles
    )
    assert 'w:styleId="PlaceholderText"' in parts["word/styles.xml"], "the main document has the style too"
    # reloaded, the controls read their entries; saved again, nothing is added twice
    again = load(io.BytesIO(saved))
    b2, r2 = _tagged(again, "c63-1", "c63-13")
    assert [b2.placeholder_text, r2.placeholder_text] == [PLACEHOLDER, "Enter a name."]
    assert len(re.findall(r"<w:docPart>", _parts(package.save())["word/glossary/document.xml"])) == 2
    assert package.last_save.added == []


def test_d9_a_glossary_of_the_document_part_alone_is_added_to_and_given_a_styles_part():
    package = load(OFFICE_JS / "65b-minimal-glossary.docx")
    package.body.content_controls[0].placeholder_text = "Something else."
    parts = _parts(package.save())
    names = re.findall(r'<w:docPartPr><w:name w:val="([^"]+)"', parts["word/glossary/document.xml"])
    assert names[:2] == [DEFAULT_ENTRY, CUSTOM_ENTRY] and len(names) == 3
    assert 'w:styleId="PlaceholderText"' in parts["word/glossary/styles.xml"]
    assert "relationships/styles" in parts["word/glossary/_rels/document.xml.rels"]
    assert package.last_save.added == ["/word/glossary/styles.xml"]


def test_d9_words_glossary_read_for_its_entries_is_saved_byte_for_byte_and_an_entry_added_leaves_its_styles_part():
    source = _parts((OFFICE_JS / "66b-word-glossary.docx").read_bytes())
    untouched = load(OFFICE_JS / "66b-word-glossary.docx")
    assert untouched.body.content_controls[0].placeholder_text == PLACEHOLDER
    assert not untouched.main_document_part.glossary_document_part.is_unmarshalled
    parts = _parts(untouched.save())
    # (a relationships part is written anew on every save)
    for name in [k for k in source if k.startswith("word/glossary/") and "_rels" not in k]:
        assert parts[name] == source[name], name
    untouched.body.content_controls[0].placeholder_text = "Another."
    added = _parts(untouched.save())
    assert added["word/glossary/styles.xml"] == source["word/glossary/styles.xml"], "it holds Placeholder Text already"
    assert "<w:t>Another.</w:t>" in added["word/glossary/document.xml"]
    assert untouched.last_save.added == []


def test_d9_a_file_with_a_glossary_never_edited_keeps_its_glossary_byte_for_byte_whatever_is_read():
    source = (ROOT / "samples" / "Normal.dotm").read_bytes()
    package = load(io.BytesIO(source))
    for body in package.bodies():
        for control in body.content_controls:
            control.placeholder_text  # noqa: B018 - read for its effect
    package.outline()
    saved = _parts(package.save())
    with zipfile.ZipFile(io.BytesIO(source)) as archive:
        for name in archive.namelist():
            if name.startswith("word/glossary/") and "_rels" not in name:
                assert saved[name].encode("utf-8") == archive.read(name), name
    assert package.last_save.added == []


def test_d9_the_content_apis_save_holds_words_form_reopened_and_read_back_and_round_trips_canonically():
    """A created document's placeholders, saved: the glossary made, read back through placeholder_text, and saved
    again unchanged in canonical form (scripts/canon.py), the save's normal form (saves must be clean)."""
    import sys

    if str(ROOT / "scripts") not in sys.path:
        sys.path.insert(0, str(ROOT / "scripts"))
    import canon as module

    package = create_package()
    body = package.body
    made = body.insert_paragraph("").insert_content_control()
    own = body.insert_paragraph("To: ").get_range("End").insert_content_control("PlainText")
    own.placeholder_text = "Enter the recipient."
    first = package.save()
    reopened = load(io.BytesIO(first))
    a, b = reopened.body.content_controls
    assert [(a.placeholder_entry, a.placeholder_text), (b.placeholder_text, b.is_showing_placeholder)] == [
        (made.placeholder_entry, PLACEHOLDER),
        ("Enter the recipient.", True),
    ]
    # unmarshalled and saved again, every glossary part is the first save's, canonically
    reopened.main_document_part.glossary_document_part.contents  # noqa: B018 - unmarshalled on purpose
    reopened.main_document_part.glossary_document_part.style_definitions_part.contents  # noqa: B018
    second = _parts(reopened.save())
    for name in ("word/glossary/document.xml", "word/glossary/styles.xml"):
        result = module.compare(_parts(first)[name].encode("utf-8"), second[name].encode("utf-8"), normalise=True)
        assert result.identical and not result.counts, (name, result.counts)


def test_a_trial_has_its_own_entries_thrown_away_with_it():
    package = _check_63()
    (block,) = _tagged(package, "c63-1")
    block.placeholder_text = "Real."
    real = package.glossary_entries.pending
    with package.dry_run() as trial:
        (copy,) = [c for c in trial.body.content_controls if c.tag == "c63-1"]
        copy.placeholder_text = "Only in the trial."
        assert trial.glossary_entries is not package.glossary_entries
        assert set(real) <= set(trial.glossary_entries.pending)
        assert len(trial.glossary_entries.pending) == len(real) + 1
    assert package.glossary_entries.pending == real
    assert "Only in the trial." not in _parts(package.save())["word/glossary/document.xml"]


# ---------------------------------------------------------------------------
# core-ts section 63.18 (aefc41f): a placeholder shows its entry's content; a missing entry gives way to Word's default
# ---------------------------------------------------------------------------


def _with_entry_of_two_paragraphs() -> tuple[Any, str]:
    """66a's glossary given an entry of two paragraphs, a bold run and a plain one; a centred block control in red and
    not bold naming it, and a run-level one (core-ts's 63.18 test)."""
    name = "6767ABCDEF0123456789ABCDEF012345"
    entry = (
        f'<w:docPart><w:docPartPr><w:name w:val="{name}"/><w:category><w:name w:val="General"/>'
        '<w:gallery w:val="placeholder"/></w:category><w:types><w:type w:val="bbPlcHdr"/></w:types>'
        '<w:behaviors><w:behavior w:val="content"/></w:behaviors></w:docPartPr><w:docPartBody>'
        '<w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr><w:t xml:space="preserve">Enter </w:t></w:r>'
        '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>the name,</w:t></w:r></w:p>'
        '<w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>then the date.</w:t></w:r></w:p>'
        "</w:docPartBody></w:docPart>"
    )
    controls = (
        '<w:sdt><w:sdtPr><w:rPr><w:b w:val="0"/><w:color w:val="FF0000"/></w:rPr><w:tag w:val="block"/>'
        f'<w:placeholder><w:docPart w:val="{name}"/></w:placeholder></w:sdtPr><w:sdtContent><w:p><w:pPr><w:jc w:val="center"/>'
        "</w:pPr><w:r><w:t>VALUE</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="run"/>'
        f'<w:placeholder><w:docPart w:val="{name}"/></w:placeholder></w:sdtPr><w:sdtContent><w:r><w:t>VALUE</w:t></w:r>'
        "</w:sdtContent></w:sdt></w:p><w:sectPr>"
    )
    out = io.BytesIO()
    with zipfile.ZipFile(OFFICE_JS / "66a-glossary-with-styles.docx") as source, zipfile.ZipFile(out, "w") as target:
        for info in source.infolist():
            data = source.read(info)
            if info.filename == "word/glossary/document.xml":
                data = data.replace(b"</w:docParts>", entry.encode() + b"</w:docParts>")
            elif info.filename == "word/document.xml":
                data = data.replace(b"<w:sectPr>", controls.encode(), 1)
            target.writestr(info, data)
    return load(io.BytesIO(out.getvalue())), name


def _bold(r_pr: Any) -> str:
    b = getattr(r_pr, "b", None)
    return "unset" if b is None else "off" if b.val is False else "on"


def _colour(r_pr: Any) -> str | None:
    value = getattr(getattr(r_pr, "color", None), "val", None)
    return value.hex().upper() if isinstance(value, bytes) else value


def test_section_63_18_the_combination_a_control_shows_a_copy_of_its_entrys_content_the_controls_run_properties_over_it():
    """The tests that pin how an entry's runs and a control's ``w:sdtPr/w:rPr`` combine, together (core-ts ``aefc41f``).

    core-ts's check 68 has since found that Word brings an entry back as the
    glossary holds it, without the control's properties over it (CR-019
    section 12); these move together when that is ported.
    """
    package, name = _with_entry_of_two_paragraphs()
    block, run_control = _tagged(package, "block", "run")
    assert block.placeholder_text == "Enter the name,\rthen the date."
    block.show_placeholder()
    run_control.show_placeholder()
    # the block control: both paragraphs, each run as the entry holds it with the control's properties over it (its
    # w:b off over the entry's bold), the second paragraph with a copy of the control's paragraph's properties
    runs = [
        [(text_of(r), r.r_pr.r_style.val, _bold(r.r_pr), _colour(r.r_pr)) for r in p.runs]
        for p in block.paragraphs
    ]
    assert runs == [
        [("Enter ", "PlaceholderText", "off", "FF0000"), ("the name,", "PlaceholderText", "off", "FF0000")],
        [("then the date.", "PlaceholderText", "off", "FF0000")],
    ]
    assert [getattr(getattr(getattr(p.element.p_pr, "jc", None), "val", None), "value", None) for p in block.paragraphs] == [
        "center",
        "center",
    ]
    assert [block.is_showing_placeholder, block.text.replace("\r", ""), block.placeholder_entry] == [True, "", name]
    # the run-level control, no run properties of its own: the first paragraph's runs as the entry holds them
    assert re.search(
        r'<w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr><w:t xml:space="preserve">Enter </w:t>'
        r'</w:r><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>the name,</w:t></w:r></w:sdtContent>',
        run_control.get_xml(),
    )
    assert [run_control.is_showing_placeholder, run_control.text, run_control.placeholder_entry] == [True, "", name]
    # the entry itself is left as it was
    assert f'<w:name w:val="{name}"/>' in _parts(package.save())["word/glossary/document.xml"]
    # the one helper the combination is made in: the control's properties over the run's, the run's alone without
    from docx4j_py.wml import RPr, el

    run_r_pr = RPr(r_style=el.rStyle(val="PlaceholderText"), b=el.b())
    combined = entry_run_properties(run_r_pr, RPr(b=el.b(val=False), color=el.color(val="FF0000")))
    assert (combined.r_style.val, _bold(combined), _colour(combined)) == ("PlaceholderText", "off", "FF0000")
    alone = entry_run_properties(run_r_pr, None)
    assert alone is not run_r_pr and (alone.r_style.val, _bold(alone), alone.color) == ("PlaceholderText", "on", None)


def test_section_63_18_a_text_control_naming_a_missing_entry_is_given_words_default_and_a_date_control_keeps_its_name():
    package = load(OFFICE_JS / "65b-minimal-glossary.docx")
    control = next(c for c in package.body.content_controls if c.placeholder_entry == MISSING_ENTRY)
    assert control.placeholder_text is None
    control.show_placeholder()
    # named the default entry, which 65b's glossary holds, and showing it; nothing written under the missing name
    assert [control.placeholder_entry, control.placeholder_text, control.is_showing_placeholder, control.text] == [
        DEFAULT_ENTRY,
        PLACEHOLDER,
        True,
        "",
    ]
    glossary = _parts(package.save())["word/glossary/document.xml"]
    assert re.findall(r'<w:docPartPr><w:name w:val="([^"]+)"', glossary) == [DEFAULT_ENTRY, CUSTOM_ENTRY]
    # Word's default entry is its text controls': a date control naming a missing entry keeps the name, and an empty
    # bound value leaves it docx4j's run (docx4j CR-063's D3)
    out = io.BytesIO()
    with zipfile.ZipFile(OFFICE_JS / "65b-minimal-glossary.docx") as source, zipfile.ZipFile(out, "w") as target:
        for info in source.infolist():
            data = source.read(info)
            if info.filename == "word/document.xml":
                data = data.replace(
                    b"<w:sectPr>",
                    b'<w:sdt><w:sdtPr><w:tag w:val="date"/><w:placeholder><w:docPart w:val="'
                    + MISSING_ENTRY.encode()
                    + b'"/></w:placeholder><w:date/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>1 January</w:t></w:r></w:p>'
                    b"</w:sdtContent></w:sdt><w:sectPr>",
                    1,
                )
            target.writestr(info, data)
    dated = load(io.BytesIO(out.getvalue()))
    (date,) = _tagged(dated, "date")
    assert [date.type, date.is_text_control, date.shows_words_form_placeholder] == ["DatePicker", False, False]
    date.show_placeholder()
    assert date.placeholder_entry == MISSING_ENTRY


def test_the_glossary_is_read_on_first_use_where_no_body_has_read_it_yet():
    """core-ts reads the glossary when a body is first read, and until then keeps a name it can't see; reads here are
    synchronous, so the entries are read on first use too, and that case does not arise (CR-019 section 12)."""
    package = load(OFFICE_JS / "65b-minimal-glossary.docx")
    entries = package.glossary_entries
    assert isinstance(entries, GlossaryEntries) and not entries.loaded
    assert entries.text_of(CUSTOM_ENTRY) == CUSTOM_TEXT and entries.loaded
    assert entries.text_of(MISSING_ENTRY) is None
    assert not package.main_document_part.glossary_document_part.is_unmarshalled
    # a body read reads them too (core-ts's getBody)
    fresh = load(OFFICE_JS / "65b-minimal-glossary.docx")
    fresh.body  # noqa: B018 - read for its effect
    assert fresh.glossary_entries.loaded
