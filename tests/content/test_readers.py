"""The readers as Office JS reads them: CR-016 Phases B and C (a placeholder is no text, a deleted mark is nothing, K1 pinned, wildcards across marks).

    .venv-fork/bin/python -m pytest tests/content/test_readers.py

The port of docx4j-core-ts's ``test/readers.test.mjs`` at ``222cc20`` for its
CR-012 sections 18 and 19 and its check 58, test for test, its strings kept:
check 58's file and snippet agreeing, its three JSONs and Word's save, the
snippet through the port held to Word 365's Simple Markup run and its Edit run
but for the differences listed (K1, K3, and case 13's two classes, K6),
a deleted mark read as nothing in Word's default view, and a control showing
its placeholder read as no text and written into in the placeholder's place.
core-ts runs check 58's Script Lab snippet through its Office JS shim; this
package has none, so ``tests/office_js/checks.py`` ports it
(:func:`run_check_58`), with the same JSON shape and ``vis()``. And, at
``3ba1e68`` (CR-016 Phase C), its check 59: the file and snippet agreeing, the
JSON and the file as copied, the snippet through the port
(:func:`run_check_59`) Word 365's run in every scope, and the refusals of a
wildcard replacement holding a mark.

Then this package's own: cases 20 to 22's markup against Word's save, a bound
control written into and its node, the text model unmoved, a placeholder's
runs without an ``office_offset``, ``^p`` and the views, a cell's and a
comment's text, a refusal leaving the placeholder, and CR-016 section 7's
risk 2 (``paragraph.text = paragraph.text`` over a placeholder is refused).
"""

from __future__ import annotations

import hashlib
import importlib.util
import io
import json
import re
from typing import Any

import pytest
from conftest import ROOT, reloaded

from docx4j_py import create_package, load
from docx4j_py.model.content import SpanError
from docx4j_py.model.text import extract, plain_text
from docx4j_py.openpackaging.load import LoadOptions
from docx4j_py.traversal import element_name
from docx4j_py.wml import to_xml

OFFICE_JS = ROOT / "tests" / "office_js"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_spec = importlib.util.spec_from_file_location("office_js_checks", OFFICE_JS / "checks.py")
checks = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(checks)

#: Check 58's cases as ``build.mjs`` at core-ts ``222cc20`` builds them (``check58Cases``): ``(n, spec)``.
CASES_58 = [
    (1, "changes"),
    (2, "changes"),
    (3, "changes"),
    (4, "changes"),
    (5, "changes"),
    (6, "range head6a body6b"),
    (7, "range two7a two7b"),
    (8, "range plain8a plain8c"),
    (9, "cell"),
    (10, "cell"),
    (11, "placeholder c58-11"),
    (12, "search placeholder"),
    (13, "search wildcards"),
    (14, "accept 1"),
    (15, "accept 2"),
    (16, "reject 1"),
    (17, "reject 2"),
    (18, "accept 2"),
    (19, "reject 2"),
    (20, "fill c58-20 replace"),
    (21, "fill c58-21 end"),
    (22, "fill c58-22 end"),
]

#: The files Word 365 read and saved, and its three JSONs, as core-ts ``222cc20`` has them (CR-016 section 6.1).
SHA1_58 = {
    "58-measures.docx": "03647705a7461f9ec6a441213f8e96d3ca69a32e",
    "58-measures-word365.docx": "0e5184c5ad383c67c68dd08d16ed26fe58e41a8c",
    "58-measures-word365-read-simple.json": "d299b8e66a56614dd5a6374a65e9130dcc655548",
    "58-measures-word365-read-all.json": "989f8a9ffef9ec98e86322949f2fecedbd021622",
    "58-measures-word365-edit.json": "b90f818773e5b2a4a4e9f0eeee57c378a66b0981",
}


def word(name: str) -> dict:
    return json.loads((OFFICE_JS / name).read_text(encoding="utf-8"))


def loaded(name: str):
    """A package of one of the files, its font selector there for ``font.name`` (core-ts's ``loaded``)."""
    package = load(OFFICE_JS / name)
    package.get_run_font_selector()
    return package


# ---------------------------------------------------------------------------
# check 58 (2026-10-23): what core-ts CR-002 section 58, CR-012 sections 18 and 19 and CR-002 section 32's
# item 1 left to measure
# ---------------------------------------------------------------------------


def test_check_58s_file_and_snippet_agree_every_case_parses_to_the_case_it_was_built_as_each_word_a_range_looks_for_and_each_tag_once():
    package = loaded("58-measures.docx")
    body = package.body
    regions = [r for p in body.paragraphs if (r := checks.parse_region_58(p.text)) is not None]
    assert [(r["n"], " ".join([r["op"], *(a.lower() if r["op"] == "range" else a for a in r["args"])])) for r in regions] == CASES_58
    ops = {"changes", "range", "cell", "placeholder", "search", "accept", "reject", "fill"}
    for r in regions:
        assert r["op"] in ops, f"case {r['n']}: {r['op']}"
        if r["op"] == "range":
            for token in r["args"]:
                assert len(body.search(token, match_case=True, match_whole_word=True)) == 1, f"case {r['n']}: {token}"
        if r["op"] in ("placeholder", "fill"):
            assert len([c for c in body.content_controls if c.tag == r["args"][0]]) == 1, f"case {r['n']}: {r['args'][0]}"
        if r["op"] == "search":
            assert callable(checks.QUERIES_58[r["args"][0]]), f"case {r['n']}"
        if r["op"] in ("accept", "reject"):
            assert r["args"][0] in ("1", "2"), f"case {r['n']}"
        if r["op"] == "fill":
            assert r["args"][1] in ("replace", "end"), f"case {r['n']}"
    # the controls' ids are distinct, and every control shows its placeholder
    ids = [c.id for c in body.content_controls]
    assert len(set(ids)) == len(ids)
    assert len(body.content_controls) == 6
    assert all(c.is_showing_placeholder for c in body.content_controls)


def test_check_58s_three_jsons_word_365_2026_10_23_are_of_this_file_and_its_save_is_words_as_copied():
    for mode, label in (("read-simple", "Simple Markup"), ("read-all", "All Markup"), ("edit", "All Markup")):
        measured = word(f"58-measures-word365-{mode}.json")
        assert [measured["label"], measured["mode"]] == [label, "read" if mode.startswith("read") else "edit"]
        expected = [n for n, spec in CASES_58 if mode != "edit" or re.match(r"^(accept|reject|fill) ", spec)]
        assert [c["n"] for c in measured["cases"]] == expected, mode
        assert measured["host"] == {"host": "Word", "version": "16.0.20326.20232", "platform": "PC"}
    for name, digest in SHA1_58.items():
        assert hashlib.sha1((OFFICE_JS / name).read_bytes()).hexdigest() == digest, name


def _listing(reads: Any) -> Any:
    """A listing as K1 pins it: the region's changes and each paragraph's, its whole range's and its content range's, by type."""
    if not isinstance(reads, dict):
        return reads

    def types(changes: Any) -> Any:
        return [c["type"] for c in changes] if isinstance(changes, list) else changes

    paragraphs = reads["paragraphs"]
    return {
        "region": types(reads["region"]),
        "paragraphs": [[types(p[k]) for k in ("changes", "wholeChanges", "contentChanges")] for p in paragraphs]
        if isinstance(paragraphs, list)
        else paragraphs,
    }


def _placeholder(reads: dict) -> dict:
    """placeholderText reads the shown runs here, where Word, with no glossary part, reads null (CR-016 Q6); and Word's
    showingPlaceholderText came back undefined, which JSON leaves out."""
    control = reads["control"]
    return {"control": {k: v for k, v in control.items() if k != "placeholderText"} if isinstance(control, dict) else control}


def _read_view(case: dict) -> Any:
    if case["op"] in ("changes", "accept", "reject"):
        return _listing(case["reads"])
    if case["op"] in ("placeholder", "fill"):
        return _placeholder(case["reads"])
    return case["reads"]


def _edit_view(case: dict) -> Any:
    """The Edit run was in All Markup, so the region's text after it isn't compared; what was resolved and what is left are."""
    reads = case["reads"]
    if case["op"] == "fill":
        return {"inserted": reads["inserted"], "after": _placeholder(reads["after"])}
    changes = reads["after"]["changes"]
    return {"resolved": reads["resolved"], "changes": [[x["type"], x["text"]] for x in changes] if isinstance(changes, list) else changes}


def run_58() -> dict[str, dict]:
    return {
        "read": checks.run_check_58(loaded("58-measures.docx"), "docx4j-python", "read"),
        "edit": checks.run_check_58(loaded("58-measures.docx"), "docx4j-python", "edit"),
    }


def test_check_58s_snippet_through_the_port_held_to_word_365s_simple_markup_run_and_its_edit_run_but_for_the_differences_listed():
    ours = run_58()
    theirs = {"read": word("58-measures-word365-read-simple.json"), "edit": word("58-measures-word365-edit.json")}
    views = {"read": _read_view, "edit": _edit_view}
    differing = []
    for mode in ("read", "edit"):
        assert [c["n"] for c in ours[mode]["cases"]] == [c["n"] for c in theirs[mode]["cases"]], mode
        for index, case in enumerate(theirs[mode]["cases"]):
            if json.dumps(views[mode](case)) != json.dumps(views[mode](ours[mode]["cases"][index])):
                differing.append(f"{mode} {case['n']}")
    # K1: Word lists nothing in two deleted paragraphs in a row (2, and 3, whose region lists only the first author's
    # deletion), nor in a paragraph whose text alone is deleted (4); this package lists each in its paragraph. K3: Word's
    # search finds the placeholder's text, reading each hit as '' (12). K6: Word's classes match nothing in the emoji or
    # the Arabic word, where [!a-z ]@ finds them here, and [!a-z ] the emoji, a code point to Python (13); its mark, as
    # Word's since CR-016 Phase C (K7)
    assert differing == ["read 2", "read 3", "read 4", "read 12", "read 13"]

    def reads(n: int) -> Any:
        return ours["read"]["cases"][n - 1]["reads"]

    assert _listing(reads(2))["paragraphs"][2:4] == [[["Deleted"], ["Deleted"], ["Deleted"]], [[], [], []]]
    assert _listing(theirs["read"]["cases"][1]["reads"])["paragraphs"][2:4] == [[[], [], []], [[], [], []]]
    assert reads(12)["Click or tap here to enter text."] == [[], [], ["Click or tap here to enter text."]]
    assert theirs["read"]["cases"][11]["reads"]["Click or tap here to enter text."] == [[""], [""], ["Click or tap here to enter text."]]
    # case 13 (CR-016 Phase C): every query Word's but the two classes. ? a code point, the pair whole (K4), the
    # Arabic letters with their fathas never cut (the grapheme filter, Q8), and the paragraph's mark (K7); a?b and x?y.
    # [!a-z ] finds the emoji, a code point here where core-ts's UTF-16 halves are dropped (section 4.9), and its mark,
    # as Word's; [!a-z ]@ the emoji and the Arabic word too (K6, core-ts's own answer), and its mark
    measured_13 = theirs["read"]["cases"][12]["reads"]
    classes = ("[!a-z ]", "[!a-z ]@")
    assert {q: v for q, v in reads(13).items() if q not in classes} == {q: v for q, v in measured_13.items() if q not in classes}
    assert reads(13)["?"] == [["x", "\\ud83d\\ude00", "y", " ", "a", "\\ud83d\\ude00", "b", " ", " ", "z", "\\u000d"]]
    assert reads(13)["[!a-z ]"] == [["\\ud83d\\ude00", "\\ud83d\\ude00", "\\u000d"]]
    assert reads(13)["[!a-z ]@"] == [["\\ud83d\\ude00", "\\ud83d\\ude00", "\\u0643\\u064e\\u062a\\u064e\\u0628\\u064e", "\\u000d"]]
    assert measured_13["[!a-z ]"] == measured_13["[!a-z ]@"] == [["\\u000d"]]
    # the deleted marks (6 to 10) are Word's, every read; the placeholder (11) every read but placeholderText,
    # which reads the runs here (Q6), and showingPlaceholderText, which Word gave undefined
    for n in (6, 7, 8, 9, 10):
        assert reads(n) == theirs["read"]["cases"][n - 1]["reads"], n
    assert reads(7)["range"]["text"] == "TWO7ATWO7B"
    assert reads(9)["cell"]["whole"] == "CELL9A\\u0009"
    assert reads(11)["control"] == {**theirs["read"]["cases"][10]["reads"]["control"], "placeholderText": "Click or tap here to enter text."}
    assert reads(11)["showingPlaceholderText"] is True and "showingPlaceholderText" not in theirs["read"]["cases"][10]["reads"]


def _control_markup(package: Any, tag: str) -> dict[str, Any]:
    """A control's properties by name, and its runs: each one's run properties by name and its text."""
    control = checks.control_by_tag(package.body, tag)
    properties = [(element_name(item) or "").rpartition("}")[2] for item in control.sdt_pr.content]
    runs = []
    for run in control._content_runs():
        r_pr = getattr(run, "r_pr", None)
        names = sorted(name for name in ("r_style", "b", "i", "color", "sz") if r_pr is not None and getattr(r_pr, name) is not None)
        runs.append((names, "".join(getattr(item, "value", "") or "" for item in run.content)))
    return {"properties": properties, "runs": runs}


def test_check_58s_edit_run_leaves_cases_20_to_22_as_words_save_after_its_runs_no_placeholder_flag_the_run_plain():
    package = loaded("58-measures.docx")
    checks.run_check_58(package, "docx4j-python", "edit")
    saved = loaded("58-measures-word365.docx")
    for n in (20, 21, 22):
        mine, theirs = _control_markup(package, f"c58-{n}"), _control_markup(saved, f"c58-{n}")
        assert mine == theirs, n
        assert theirs == {"properties": ["alias", "tag", "id"], "runs": [([], f"FILL{n}")]}, n
    # what the read run reads of Word's save is what this package's edit leaves
    for n in (20, 21, 22):
        control = checks.control_by_tag(saved.body, f"c58-{n}")
        assert (control.text, control.is_showing_placeholder) == (f"FILL{n}", False)


# ---------------------------------------------------------------------------
# check 59 (2026-10-23): where Word's wildcards match a paragraph mark (core-ts CR-002 section 32's K5, K7)
# ---------------------------------------------------------------------------

#: Check 59's cases as ``build.mjs`` at core-ts ``222cc20`` builds them (``check59Cases``): ``(n, shape)``.
CASES_59 = [(1, "body"), (2, "cell-only"), (3, "cell-two"), (4, "cells-rows"), (5, "control")]

#: Each case's content as ``build.mjs`` builds it: the region's paragraphs between BEGIN and END, each table's
#: cells' paragraphs by row, each control's tag, id, alias and paragraphs.
CONTENT_59 = {
    1: (["Q", "Z"], [], []),
    2: (["Q", "Z"], [[[["Q"]]]], []),
    3: (["Q", "Z", "W"], [[[["Q", "Z"]]]], []),
    4: (["Q", "Z", "W", "X", "Y"], [[[["Q"], ["Z"]], [["W"], ["X"]]]], []),
    5: (["Q", "Z", "W"], [], [("c59-59000", 59000, "Check 59", ["Q", "Z"])]),
}

#: The file Word 365 read and its JSON, as CR-016 section 6.1 has them (the JSON core-ts ``222cc20``'s, unchanged at ``3ba1e68``).
SHA1_59 = {
    "59-wildcard-marks.docx": "30e8d448f10639f6fa27fb0bb7c00c920ac70117",
    "59-wildcard-marks-word365.json": "400556cd1cdb419f70ed448b6ff15f04f77bd20d",
}


def test_check_59s_file_and_snippet_agree_every_case_parses_to_the_shape_it_was_built_as_each_token_once():
    body = load(OFFICE_JS / "59-wildcard-marks.docx").body
    regions = [r for p in body.paragraphs if (r := checks.parse_region_59(p.text)) is not None]
    assert [(r["n"], r["shape"]) for r in regions] == CASES_59
    for n, _shape in CASES_59:
        for token in (f"BEGIN{n}", f"END{n}"):
            assert len(body.search(token, match_case=True, match_whole_word=True)) == 1, token
        region = checks.region_of_59(body, n)
        tables = [
            [[[p.text for p in table.cell(i, j).body.paragraphs] for j in range(len(row.cells))] for i, row in enumerate(table.rows)]
            for table in region.tables
        ]
        controls = [(c.tag, c.id, c.title, [p.text for p in c.body().paragraphs]) for c in region.content_controls]
        assert ([p.text for p in region.paragraphs[1:-1]], tables, controls) == CONTENT_59[n], n
    # the queries are the snippet's, and its one control's tag is there once
    assert checks.REGION_QUERIES_59 == ("Q?Z", "Q^13Z", "Q*Z", "^13")
    assert checks.PART_QUERIES_59 == ("?", "[!A-Z]", "^13")
    assert [c.tag for c in body.content_controls] == ["c59-59000"]


def test_check_59s_json_word_365_2026_10_23_is_of_this_file_and_the_files_are_as_copied():
    measured = word("59-wildcard-marks-word365.json")
    assert [measured["check"], measured["label"]] == [59, "Simple Markup"]
    assert measured["host"] == {"host": "Word", "version": "16.0.20326.20232", "platform": "PC"}
    assert [(c["n"], c["shape"]) for c in measured["cases"]] == CASES_59
    for name, digest in SHA1_59.items():
        assert hashlib.sha1((OFFICE_JS / name).read_bytes()).hexdigest() == digest, name


def test_check_59s_snippet_through_the_port_is_word_365s_run_wildcards_match_a_mark_where_p_does_k7():
    """core-ts: "check 59's file and snippet agree, and the snippet through the Word shim is Word 365's run: wildcards
    match a mark where ^p does (K7)", the agreement test above and this one."""
    package = load(OFFICE_JS / "59-wildcard-marks.docx")
    body = package.body
    result = checks.run_check_59(package, "docx4j-python")
    assert len(result["cases"]) == 5
    assert "error: " not in json.dumps(result)
    # Word 365's run (2026-10-23), scope for scope: a paragraph's mark matched, never a cell's or a row's end, nor across a
    # table's edge, as `^p`'s; a control's range without its last paragraph's mark
    measured = word("59-wildcard-marks-word365.json")
    assert [c["shape"] for c in measured["cases"]] == [c["shape"] for c in result["cases"]]
    for index, case in enumerate(measured["cases"]):
        assert result["cases"][index]["reads"] == case["reads"], case["shape"]
    assert measured["cases"][1]["reads"]["cell"]["?"] == ["Q"], "a cell's end: no mark"
    assert measured["cases"][3]["reads"]["table"]["?"] == ["Q", "Z", "W", "X"], "cells' and rows' ends: no mark"
    assert measured["cases"][4]["reads"]["control"]["?"] == ["Q", "\\u000d", "Z"], (
        "inside a control: its paragraphs' marks but the last, which its range doesn't hold"
    )
    # a replacement where a hit holds a mark is refused before anything changes, as a plain ^p's (Q9: SpanError
    # range.replaces_mark, where core-ts throws NotSupportedError)
    before = body.get_xml()
    with pytest.raises(SpanError, match="holds a paragraph mark") as refused:
        body.replace_text("Q?Z", "X", match_wildcards=True)
    assert refused.value.code == "range.replaces_mark"
    with pytest.raises(SpanError, match="holds a paragraph mark") as refused:
        body.paragraphs[2].replace_text("?", "X", match_wildcards=True)
    assert refused.value.code == "range.replaces_mark"
    assert body.get_xml() == before
    paragraph = body.paragraphs[2]
    assert [paragraph.text, paragraph.replace_text("Q", "X", match_wildcards=True), paragraph.text] == ["Q", 1, "X"], (
        "a hit holding no mark is replaced"
    )


# ---------------------------------------------------------------------------
# CR-012 sections 18 and 19 (core-ts readers.test.mjs at 222cc20)
# ---------------------------------------------------------------------------

DELETED = '<w:del w:id="1" w:author="A" w:date="2026-10-23T00:00:00Z"/>'


def _section_19(mark: str = DELETED) -> Any:
    package = create_package()
    body = package.body
    body.insert_xml(
        f'<w:p><w:pPr><w:rPr>{mark}</w:rPr></w:pPr><w:r><w:t>ONE</w:t></w:r></w:p><w:p><w:r><w:t>TWO</w:t></w:r></w:p>'
        f'<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr><w:rPr>{mark.replace(chr(34) + "1" + chr(34), chr(34) + "2" + chr(34))}</w:rPr></w:pPr><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        "<w:p><w:r><w:t>LAST</w:t></w:r></w:p>",
        location="End",
    )
    return package


def test_cr_012_section_19_a_deleted_mark_reads_as_nothing_in_words_default_view_the_paragraphs_apart_a_cells_end_kept_check_58():
    body = _section_19().body
    one, two = body.paragraphs[0], body.paragraphs[1]
    assert one.mark_text == ""
    assert one.get_range("Whole").text == "ONE"
    assert one.get_range("Whole").expand_to(two.get_range("Whole")).text == "ONETWO\r"
    assert len(body.paragraphs) == 4, "the paragraphs stay apart"
    # a cell's last paragraph ends with the cell's \t whatever its mark (check 58 cases 9 and 10)
    assert body.paragraphs[2].mark_text == "\t"
    assert body.text == "ONETWO\rCELL\r\nLAST"
    assert body.get_text(view="original") == "ONE\rTWO\rCELL\r\nLAST", "the original view keeps a deleted mark"


def test_the_original_view_writes_no_inserted_mark_and_the_default_view_keeps_it_the_marks_reject_all_and_accept_all_keep():
    inserted = DELETED.replace("w:del ", "w:ins ")
    body = _section_19(inserted).body
    assert body.text == "ONE\rTWO\rCELL\r\nLAST"
    assert body.get_text(view="original") == "ONETWO\rCELL\r\nLAST"
    assert body.paragraphs[0].mark_text == "\r"
    assert body.paragraphs[2].mark_text == "\t"


def test_p_stops_matching_a_deleted_mark_and_a_cells_text_a_comments_and_the_text_model_read_as_before():
    package = _section_19()
    body = package.body
    # ^p: TWO's mark and LAST's (the body's Whole range holds it); ONE's deleted mark is nothing, CELL's a cell's end
    assert [r.text for r in body.search("^p")] == ["\r", "\r"]
    assert body.paragraphs[0].search("^p") == []
    assert [r.text for r in body.paragraphs[1].search("^p")] == ["\r"]
    assert body.tables[0].values == [["CELL"]]
    # the text model reads each view as its resolution leaves it, as before (CR-016 section 4.8 item 4)
    assert plain_text(extract(package)) == "ONETWO\nCELL\nLAST\n"
    assert plain_text(extract(package, view="original")) == "ONE\nTWO\nCELL\nLAST\n"
    # a deleted mark's change keeps its text, a mark's \r: CR-016 section 4.5, TrackedChange.text not touched
    assert [(c.type, c.text) for c in body.get_tracked_changes()] == [("Deleted", "\r"), ("Deleted", "\r")]


def test_a_comment_or_a_block_control_holding_a_deleted_mark_reads_its_paragraphs_joined_there():
    package = create_package()
    body = package.body
    body.insert_xml(
        f"<w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr><w:sdtContent><w:p><w:pPr><w:rPr>{DELETED}</w:rPr></w:pPr>"
        "<w:r><w:t>IN1</w:t></w:r></w:p><w:p><w:r><w:t>IN2</w:t></w:r></w:p></w:sdtContent></w:sdt>",
        location="End",
    )
    control = body.content_controls[0]
    assert control.text == "IN1IN2"
    assert [p.text for p in control.paragraphs] == ["IN1", "IN2"]
    assert control.get_range("Whole").text == "IN1IN2\r"


PLACEHOLDER = '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click here.</w:t></w:r>'


def _sdt_pr(identifier: int, extra: str = "") -> str:
    return f'<w:sdtPr><w:id w:val="{identifier}"/>{extra}<w:showingPlcHdr/></w:sdtPr>'


def _section_18() -> Any:
    package = create_package()
    package.body.insert_xml(
        f"<w:sdt>{_sdt_pr(1)}<w:sdtContent><w:p>{PLACEHOLDER}</w:p></w:sdtContent></w:sdt>"
        f'<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:sdt>{_sdt_pr(2)}<w:sdtContent>{PLACEHOLDER}</w:sdtContent></w:sdt>'
        '<w:r><w:t xml:space="preserve"> B</w:t></w:r></w:p>',
        location="End",
    )
    return package


def test_cr_012_section_18_a_control_showing_its_placeholder_reads_as_no_text_is_no_text_to_search_and_takes_text_in_the_placeholders_place_check_58():
    package = _section_18()
    body = package.body
    block, run = body.content_controls
    assert [block.text, run.text, body.paragraphs[1].text] == ["", "", "A  B"]
    assert [block.placeholder_text, run.placeholder_text] == ["Click here.", "Click here."], "read from the runs (Word: null, from the glossary)"
    assert len(body.search("Click")) == 0, "K3: no text to search"
    assert body.text == "\rA  B"
    # text goes in the placeholder's place, at 'End' as at 'Replace'; the control shows it, not the placeholder style
    block.insert_text("FILLED", location="End")
    run.insert_text("X", location="Start")
    assert [block.text, block.is_showing_placeholder, run.text, run.is_showing_placeholder] == ["FILLED", False, "X", False]
    assert body.paragraphs[1].text == "A X B"
    runs = [*block.paragraphs[0].runs, *body.paragraphs[1].runs]
    assert all(getattr(getattr(r.r_pr, "r_style", None), "val", None) != "PlaceholderText" for r in runs if r.r_pr is not None)


@pytest.mark.parametrize("location", ["Start", "End", "Replace"])
@pytest.mark.parametrize("form", ["block", "run"])
def test_every_location_writes_in_the_placeholders_place_the_flag_cleared_the_style_gone_the_runs_own_look_kept(form, location):
    package = create_package()
    look = '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr><w:t>Click here.</w:t></w:r>'
    if form == "block":
        package.body.insert_xml(f"<w:sdt>{_sdt_pr(1)}<w:sdtContent><w:p>{look}</w:p></w:sdtContent></w:sdt>", location="End")
    else:
        package.body.insert_xml(f"<w:p><w:r><w:t>A</w:t></w:r><w:sdt>{_sdt_pr(2)}<w:sdtContent>{look}</w:sdtContent></w:sdt></w:p>", location="End")
    control = package.body.content_controls[0]
    returned = control.insert_text("VALUE", location=location)
    assert returned.text == "VALUE"
    assert (control.text, control.is_showing_placeholder, control.placeholder_text) == ("VALUE", False, "")
    xml = to_xml(control.element)
    assert "showingPlcHdr" not in xml and "PlaceholderText" not in xml and "Click here." not in xml
    if form == "run":
        # the placeholder's run takes the text, its own properties but the style kept: the run's own look
        assert "<w:b/>" in xml
    again = load(io.BytesIO(package.save()), options=LoadOptions(strict=True)).body.content_controls[0]
    assert (again.text, again.is_showing_placeholder) == ("VALUE", False)


def test_a_refused_write_leaves_the_control_showing_its_placeholder():
    package = _section_18()
    run = package.body.content_controls[1]
    with pytest.raises(SpanError) as raised:
        run.insert_text("one\ntwo", location="End")
    assert raised.value.code == "range.splits_control"
    assert run.is_showing_placeholder and run.placeholder_text == "Click here."
    assert "PlaceholderText" in to_xml(run.element)


def test_a_bound_control_showing_its_placeholder_written_into_holds_a_value_its_node_takes_cr_016_section_2_4():
    package = create_package()
    package.body.insert_xml(
        '<w:p><w:r><w:t xml:space="preserve">To: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="5"/><w:showingPlcHdr/><w:text/></w:sdtPr>'
        f"<w:sdtContent>{PLACEHOLDER}</w:sdtContent></w:sdt></w:p>",
        location="End",
    )
    control = package.body.content_controls[0]
    part = package.custom_xml_parts.add('<greeting xmlns="http://example.com/g"><to></to></greeting>')
    control.xml_mapping.set_mapping("/ns0:greeting[1]/ns0:to[1]", "xmlns:ns0='http://example.com/g'", part)
    node = part.select_single_node("/ns0:greeting[1]/ns0:to[1]")
    assert control.text == "" and node.text in ("", None)
    control.insert_text("VALUE", location="Replace")
    assert (control.text, control.is_showing_placeholder) == ("VALUE", False)
    result = package.custom_xml_parts.update_from_content_controls()
    assert part.select_single_node("/ns0:greeting[1]/ns0:to[1]").text == "VALUE"
    assert result is not None
    again = reloaded(package)
    assert again.custom_xml_parts.items[0].select_single_node("/ns0:greeting[1]/ns0:to[1]").text == "VALUE"
    assert again.body.content_controls[0].text == "VALUE"


def test_apply_bindings_writes_docx4js_placeholder_for_an_empty_value_which_reads_as_no_text():
    package = create_package()
    package.body.insert_xml(
        '<w:p><w:sdt><w:sdtPr><w:id w:val="6"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>old</w:t></w:r></w:sdtContent></w:sdt></w:p>',
        location="End",
    )
    control = package.body.content_controls[0]
    part = package.custom_xml_parts.add('<greeting xmlns="http://example.com/g"><to></to></greeting>')
    control.xml_mapping.set_mapping("/ns0:greeting[1]/ns0:to[1]", "xmlns:ns0='http://example.com/g'", part)
    package.custom_xml_parts.apply_bindings()
    assert control.is_showing_placeholder
    assert control.text == ""
    assert control.placeholder_text == "Click here to enter text."


def test_the_text_model_reads_a_placeholder_as_text_and_its_runs_have_no_office_offset():
    package = _section_18()
    model = extract(package)
    assert plain_text(model) == "Click here.\nA Click here. B\n"
    runs = [run for block in model.blocks for run in block.runs]
    placeholder = [run for run in runs if run.text == "Click here."]
    assert len(placeholder) == 2 and all(run.office_offset is None for run in placeholder)
    # the runs after it in its paragraph stand where Paragraph.text has them ("A  B"): until CR-016 Phase B
    # the placeholder's at 0 and 2, and " B" at 13
    others = [run for run in runs if run.text in ("A ", " B")]
    assert [run.office_offset for run in others] == [0, 2]


def test_risk_2_paragraph_text_written_over_a_run_level_placeholder_takes_the_control_as_words_replace_does():
    # CR-016 Phase B refused the write (range.holds_placeholder, the user's decision of 2026-10-24, pending a
    # measurement); Word's Replace takes the control (core-ts check 63 cases 19, 26, 45 and 52), and so does the
    # setter since CR-019 (section 4.3, core-ts's D3): untracked it goes with the content, tracked it is marked
    # deleted in Word's form, between w:customXmlDelRange markers, nothing inside marked deleted, its placeholder gone
    # unmarked and its flag cleared (CR-019 Phase A2, core-ts's D3 at d84f3da, check 63's run 2; Phase A put the
    # placeholder in a w:del inside and kept the flag, as run 1, in Design Mode, had it)
    package = _section_18()
    paragraph = package.body.paragraphs[1]
    paragraph.text = paragraph.text
    assert (paragraph.text, len(package.body.content_controls)) == ("A  B", 1)
    assert "w:sdt" not in to_xml(paragraph.element)
    package = _section_18()
    package.change_tracking_mode = "TrackAll"
    paragraph = package.body.paragraphs[1]
    paragraph.text = "C"
    xml = to_xml(paragraph.element)
    assert paragraph.text == "C"
    assert re.search(
        r'<w:customXmlDelRangeStart [^>]*/><w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent>'
        r"<w:customXmlDelRangeEnd [^>]*/><w:customXmlDelRangeStart [^>]*/></w:sdtContent></w:sdt>"
        r"<w:customXmlDelRangeEnd [^>]*/>",
        xml,
    ), xml
    assert "Click here." not in xml
    # once written into, the control holds a value and the paragraph's text is the control's too
    package = _section_18()
    package.body.content_controls[1].insert_text("X")
    paragraph = package.body.paragraphs[1]
    paragraph.text = paragraph.text
    assert paragraph.text == "A X B"


def test_range_font_a_read_splits_no_run_and_a_write_formats_the_span_alone_not_the_rest_of_its_run():
    package = create_package()
    package.body.insert_xml("<w:p><w:r><w:t>flat</w:t><w:tab/><w:t>really</w:t></w:r></w:p>", location="End")
    p = package.body.paragraphs[-1]
    before = to_xml(p.element)
    assert p.search("lat")[0].font.bold is False
    assert to_xml(p.element) == before, "a read changed nothing"
    p.search("lat")[0].font.bold = True
    xml = to_xml(p.element)
    assert re.search(
        r"<w:r><w:t>f</w:t></w:r><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>lat</w:t></w:r><w:r><w:tab/><w:t>really</w:t></w:r>",
        xml,
    ), xml
