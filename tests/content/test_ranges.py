"""Ranges across paragraphs, as Office JS's ``Word.Range`` is: CR-013 Phases B and C (the reads and the writes).

    .venv-fork/bin/python -m pytest tests/content/test_ranges.py

The port of docx4j-core-ts's ``test/ranges.test.mjs`` at ``202a01d`` (its
CR-014), test for test, its strings kept: the five of phase 0 (check 52's file
and snippet agreeing, the ``TABLE_OOXML`` package, the two JSONs of this file,
Word's saves, ``accept_all`` on Word's save after the run against Word's Accept
All), the five of phase 1 (check 52's ``reads`` and ``report`` modes against
Word's, the tracked changes a range lists, a range from ``expand_to`` and the
views across paragraphs), the two of check 53, phase 2's writes (check 52's
``run`` mode against Word's run, its Accept All and its markup) and phase 3's
(check 54's, the same three ways, with its fixtures' manifest), as CR-013 Phase
C ported them. core-ts runs checks 52's, 53's and 54's Script Lab snippets
through its Office JS shim; this package has none, so
``tests/office_js/checks.py`` ports them onto the content API, with the same
JSON shape and ``vis()`` (CR-013 sections 4.9 item 6 and 6.1).

Every difference from Word is asserted as it stands, each with its reason, as
core-ts asserts them; where this package's rule differs from core-ts's the test
asserts this package's answer and names the item (CR-013 section 4.9). Then
Phase B's own: the refusals of every write across paragraphs, before anything
changes, Word's own refusals inside a table among them, the constructor's
checks, the comment's extent over the story (section 2.5 item 1), check 48's
spans and the members the subset adds.
"""

from __future__ import annotations

import dataclasses
import functools
import hashlib
import importlib.util
import io
import json
import re
from typing import Any

import pytest
from conftest import ROOT

from docx4j_py import create_package, load
from docx4j_py.model import revisions
from docx4j_py.model.content import ContentError, MixedFont, Range, SpanError
from docx4j_py.traversal import element_name

OFFICE_JS = ROOT / "tests" / "office_js"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_spec = importlib.util.spec_from_file_location("office_js_checks", OFFICE_JS / "checks.py")
checks = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(checks)


def word(name: str) -> dict:
    return json.loads((OFFICE_JS / name).read_text(encoding="utf-8"))


def loaded(name: str):
    """A package of one of the files, as core-ts's ``loaded``: its font selector there for ``font.name``."""
    package = load(OFFICE_JS / name)
    package.get_run_font_selector()
    return package


def word_in(package, token: str) -> Range:
    """The file's paragraph holding a word, and a range over that word (core-ts's ``wordIn``)."""
    return package.body.search(token, match_case=True, match_whole_word=True)[0]


def flat(value: Any, prefix: str = "", out: dict | None = None) -> dict:
    """A JSON value as ``path -> leaf``, as core-ts's test flattens check 53's reads."""
    out = {} if out is None else out
    if isinstance(value, dict):
        for key, item in value.items():
            flat(item, f"{prefix}.{key}" if prefix else key, out)
    elif isinstance(value, list):
        for index, item in enumerate(value):
            flat(item, f"{prefix}.{index}" if prefix else str(index), out)
    else:
        out[prefix] = value
    return out


# ---------------------------------------------------------------------------
# check 52's cases, as core-ts's build.mjs (8fdccf8) lists them
# ---------------------------------------------------------------------------

_ALL = ["HB", "BH", "LB", "BL", "HMB", "TBL", "INTO", "CELL", "HID", "FN", "CC", "W1", "W2"]
_MATRIX = [
    ("delete", _ALL, ("untracked", "tracked")),
    ("replace", ["HB", "BH", "LB", "BL", "TBL", "INTO", "CELL", "CC", "W1", "W2"], ("untracked", "tracked")),
    ("bold", ["HB", "TBL", "CELL", "CC"], ("untracked", "tracked")),
    ("paragraph-before", ["HB", "TBL", "CELL", "CC", "W1", "W2"], ("untracked",)),
    ("paragraph-after", ["HB", "TBL", "CELL", "CC", "W1", "W2"], ("untracked",)),
    ("paragraph-before", ["HB"], ("tracked",)),
    ("paragraph-after", ["HB"], ("tracked",)),
    ("text-start", ["HB", "CC", "W1", "W2"], ("untracked",)),
    ("text-end", ["HB", "CC", "W1", "W2"], ("untracked",)),
    ("comment", ["HB", "TBL", "CELL", "CC", "W1", "W2", "FN"], ("untracked",)),
    ("table-end", ["TC1"], ("untracked", "tracked")),
    ("paragraph-delete", ["TC1", "TC2"], ("untracked", "tracked")),
    ("clear", ["TC1"], ("untracked", "tracked")),
    ("ooxml-replace", ["TC1"], ("untracked",)),
    ("control-delete", ["TC3"], ("untracked",)),
]


def check_52_cases() -> list[tuple[int, str, str, bool]]:
    """``check52Cases()``: (n, kind, op, tracked), in document order."""
    cases = [(1, "VIEWS", "read", False), (2, "SEARCH", "search", False)]
    for op, kinds, trackings in _MATRIX:
        for tracking in trackings:
            for kind in kinds:
                cases.append((len(cases) + 1, kind, op, tracking == "tracked"))
    return cases


# ---------------------------------------------------------------------------
# phase 0: check 52's file, snippet and Word's answers (core-ts 8fdccf8)
# ---------------------------------------------------------------------------


def test_check_52s_file_and_snippet_agree_every_case_parses_every_token_the_snippet_looks_for_is_there_once():
    package = load(OFFICE_JS / "52-ranges.docx")
    texts = [p.text for p in package.body.paragraphs]
    regions = [r for r in map(checks.parse_region_52, texts) if r is not None]
    cases = check_52_cases()
    assert len(cases) == 95
    assert [(r["n"], r["kind"], r["op"], r["tracked"]) for r in regions] == cases

    def count(token: str) -> int:
        pattern = re.compile(rf"(?<![A-Za-z0-9]){re.escape(token)}(?![A-Za-z0-9])")
        return sum(len(pattern.findall(text)) for text in texts)

    tags = [c.tag for c in package.body.content_controls]
    for r in regions:
        tokens = [f"BEGIN{r['n']}", f"END{r['n']}"]
        if "range" in r:
            spec = r["range"]
            tokens += [spec[k] for k in ("from", "to", "toMark", "whole", "cell", "paragraph") if k in spec]
            if "control" in spec:
                assert tags.count(spec["control"]) == 1, spec["control"]
        elif "queries" in r:
            for query in r["queries"]:
                tokens += [s for s in re.split(r"\^.", query) if s]
        else:
            assert r["kind"] == "VIEWS"
        for token in tokens:
            assert count(token) == 1, f"case {r['n']}: {token}"
    # every kind of range, every operation and both tracking modes are there
    assert sorted({"+".join(r["range"]) for r in regions if "range" in r}) == [
        "cell",
        "control",
        "from+to",
        "from+toMark",
        "paragraph",
        "whole",
    ]
    assert sorted({r["op"] for r in regions}) == [
        "bold",
        "clear",
        "comment",
        "control-delete",
        "delete",
        "ooxml-replace",
        "paragraph-after",
        "paragraph-before",
        "paragraph-delete",
        "read",
        "replace",
        "search",
        "table-end",
        "text-end",
        "text-start",
    ]
    # CR-002 section 55's cases: the token each names is in a table cell, the control's in one too
    for r in (x for x in regions if x["kind"].startswith("TC")):
        spec = r["range"]
        token = spec.get("cell") or spec.get("paragraph")
        if token is not None:
            paragraph = next(p for p in package.body.paragraphs if token in p.text)
        else:
            paragraph = next(c for c in package.body.content_controls if c.tag == spec["control"]).paragraphs[0]
        assert paragraph.parent_table_cell is not None, f"case {r['n']}"


def test_check_52s_table_for_ooxml_replace_is_a_flat_opc_package_this_package_reads_a_table_of_one_cell_nothing_after_it():
    package = create_package()
    body = package.body
    body.insert_paragraph("before", location="End")
    inserted = body.insert_ooxml(checks.TABLE_OOXML, location="End")
    assert len(inserted) == 1
    assert [t.values for t in body.tables] == [[["NEWOOXML"]]]


def test_check_52s_run_and_its_report_after_accept_all_are_of_this_file_the_same_95_cases_each_its_kind_operation_and_tracking():
    for name, mode in (("52-ranges-word365-run.json", "run"), ("52-ranges-word365-accepted.json", "report")):
        measured = word(name)
        assert measured["mode"] == mode
        assert [(c["n"], c["kind"], c["op"], c["tracked"]) for c in measured["cases"]] == check_52_cases()


def test_check_52s_word_saves_are_words_as_copied_both_files_listed_each_sha1_matching():
    manifest = {
        name: digest
        for digest, name in (
            line.split("  ", 1) for line in (OFFICE_JS / "SHA1SUMS").read_text(encoding="utf-8").splitlines() if line
        )
    }
    # core-ts 8fdccf8's test/check52/manifest.json
    expected = {
        "52-ranges-word365.docx": "15f505306f8faf95f6946b47789849ccbf34603b",
        "52-ranges-word365-accepted.docx": "3d2a37367822fc323c0905e30b0931c6a3a4f58d",
    }
    for name, digest in expected.items():
        assert manifest[name] == digest
        assert hashlib.sha1((OFFICE_JS / name).read_bytes()).hexdigest() == digest, name


def regions_of(package) -> dict[int, list[tuple[str, str, str, bool, int]]]:
    """Each case's paragraphs between its BEGIN and END paragraphs: text, style, alignment, list, table nesting."""
    regions: dict[int, list] = {}
    n = None
    for p in package.body.paragraphs:
        begin = re.match(r"^BEGIN(\d+) ", p.text)
        if begin:
            n = int(begin.group(1))
            regions[n] = []
            continue
        if n is not None and p.text == f"END{n}":
            n = None
            continue
        if n is not None:
            regions[n].append((p.text, p.style_built_in, p.alignment, p.is_list_item, p.table_nesting_level))
    return regions


def test_check_52_accept_all_on_word_365s_save_after_the_run_gives_words_accept_all_region_for_region():
    ours = load(OFFICE_JS / "52-ranges-word365.docx")
    assert len(revisions.tracked_changes(ours)) > 0
    revisions.accept_all(ours)
    assert len(revisions.tracked_changes(ours)) == 0
    mine = regions_of(ours)
    theirs = regions_of(load(OFFICE_JS / "52-ranges-word365-accepted.docx"))
    assert len(theirs) == 95
    assert list(mine) == list(theirs)
    differing = [n for n in theirs if mine[n] != theirs[n]]
    # case 26 (CC, its content deleted under tracking): Word fills the emptied control with five spaces,
    # this package leaves it empty; the paragraph is the same otherwise
    assert differing == [26]
    assert [p[0] for p in theirs[26]] == ["     "]
    assert mine[26] == [("", *p[1:]) for p in theirs[26]]
    # the joins: the joined paragraph in the properties of the paragraph the range ended in (HB, BH, LB,
    # BL, TBL). Word's 'Left' is this package's "Unknown" where no w:jc says so anywhere: CR-002 section
    # 15's recorded departure (Alignment keeps "Unknown")
    assert [[p[:4] for p in theirs[n]] for n in (16, 17, 18, 19, 21)] == [
        [("HEAD16A  BODY16B", "Normal", "Centered", False)],
        [("BODY17A  HEAD17B", "Heading1", "Unknown", False)],
        [("LIST18A  BODY18B", "Normal", "Centered", False)],
        [("BODY19A  LIST19B", "ListParagraph", "Unknown", True)],
        [("HEAD21A  BODY21B", "Normal", "Centered", False)],
    ]


# ---------------------------------------------------------------------------
# phase 1, the reads (core-ts ef57729)
# ---------------------------------------------------------------------------


def test_cr_014_phase_1_check_52s_reads_run_through_the_port_on_the_same_file_are_word_365s_kind_for_kind():
    ours = checks.run_check_52(loaded("52-ranges.docx"), "reads")
    measured = word("52-ranges-word365-run.json")
    assert list(ours["reads"]) == list(measured["reads"])
    assert len(measured["reads"]) == 16
    for kind, theirs in measured["reads"].items():
        # `marks`, the hits of ^p within the range, since CR-013 Phase C (core-ts's phase 3) too: every
        # paragraph mark outside a table the range holds, never a cell's end (CELL and the TC kinds find none)
        mine = dict(ours["reads"][kind])
        theirs = dict(theirs)
        # Word gives the style 'Other', named '', to a paragraph after a nested table in a cell whose
        # style isn't Normal (TC2; check 53, cases 10 and 12), which is not followed: core-ts CR-002
        # section 57's departure, taken here (Q6). FN's 'Other', a range whose runs' character styles
        # differ, is Word's rule and followed (check 53, cases 5 and 8)
        if kind == "TC2":
            assert [theirs["range"]["style"], theirs["range"]["styleBuiltIn"]] == ["", "Other"], kind
            assert [mine["range"]["style"], mine["range"]["styleBuiltIn"]] == ["List Paragraph", "ListParagraph"]
            mine["range"] = {**mine["range"], "style": "", "styleBuiltIn": "Other"}
        assert mine == theirs, kind


#: The paragraphs check 52's report reads, but for the paragraph's font: a paragraph's font reads its first
#: run here (CR-002), where Office JS reads null where its runs differ.
def _without_font(entry: dict) -> dict:
    return {k: v for k, v in entry.items() if k not in ("bold", "italic", "underline", "words")}


def _alignment_as_ours(expected: list[dict], mine: list[dict], n: int) -> int:
    """Word's ``'Left'`` read as this package's ``"Unknown"`` where it reads that: CR-002 section 15's departure."""
    moved = 0
    for k, entry in enumerate(expected):
        if entry.get("alignment") == "Left" and k < len(mine) and mine[k].get("alignment") == "Unknown":
            entry["alignment"] = "Unknown"
            moved += 1
    return moved


def test_cr_014_phase_1_check_52s_report_run_through_the_port_on_words_save_after_accept_all_is_words_region_for_region():
    ours = checks.run_check_52(loaded("52-ranges-word365-accepted.docx"), "report")
    measured = word("52-ranges-word365-accepted.json")
    assert len(ours["cases"]) == len(measured["cases"]) == 95
    bold = {c["n"] for c in measured["cases"] if c["op"] == "bold"}
    unknown = 0
    for index, theirs in enumerate(measured["cases"]):
        n = theirs["n"]
        mine = json.loads(json.dumps(ours["cases"][index]["region"]))
        expected = json.loads(json.dumps(theirs["region"]))
        # Range.footnotes is not implemented
        assert re.match(r"^error: NotImplemented", mine["footnotes"]), n
        del expected["footnotes"], mine["footnotes"]
        # getTextRanges, which the bold cases' report reads each word with, is not implemented
        if n in bold:
            assert re.match(r"^error: NotImplemented", mine["paragraphs"]), n
            del expected["paragraphs"], mine["paragraphs"]
        else:
            expected["paragraphs"] = [_without_font(p) for p in expected["paragraphs"]]
            mine["paragraphs"] = [_without_font(p) for p in mine["paragraphs"]]
            unknown += _alignment_as_ours(expected["paragraphs"], mine["paragraphs"], n)
        # a paragraph after a nested table in a cell: Word's style 'Other' again (cases 89, 91, 94), Q6
        for k, p in enumerate(expected.get("paragraphs", [])):
            if p["styleBuiltIn"] == "Other":
                assert n in (89, 91, 94), n
                assert [mine["paragraphs"][k]["styleBuiltIn"], mine["paragraphs"][k]["style"]] == [
                    "ListParagraph",
                    "List Paragraph",
                ], n
                expected["paragraphs"][k] = {**p, "style": "List Paragraph", "styleBuiltIn": "ListParagraph"}
        # cases 13 and 26, a control emptied: Word marks it w:showingPlcHdr over a run of five spaces, and
        # reads no text there, where this package reads the spaces (core-ts CR-014 section 9.4)
        if n in (13, 26):
            expected["text"] = expected["text"].replace("\\u000d\\u000d", "\\u000d     \\u000d")
            expected["paragraphs"][1]["text"] = "     "
            expected["contentControls"][0]["text"] = "     "
        assert mine == expected, f"case {n} ({theirs['kind']} {theirs['op']}{' tracked' if theirs['tracked'] else ''})"
    # this package's Alignment reads "Unknown" where no w:jc says otherwise (CR-002 section 15), each such
    # paragraph Word's 'Left'
    assert unknown > 0


def test_cr_014_phase_1_the_tracked_changes_a_range_across_paragraphs_lists_on_words_save_after_the_run_are_words_but_for_the_listings_own_grouping():
    ours = checks.run_check_52(loaded("52-ranges-word365.docx"), "report")
    measured = word("52-ranges-word365-run.json")

    # a deletion's text and range are All Markup's in Word's run (deleted text shown), so a deletion is
    # compared by its type
    def view(changes: list[dict]) -> list[list[str]]:
        return [[c["type"], "" if c["type"] == "Deleted" else c["range"]] for c in changes]

    replaced = {39, 40, 41, 42, 43, 46, 47, 48}
    compared = []
    for index, theirs in enumerate(measured["cases"]):
        expected = theirs["region"]["trackedChanges"]
        if not isinstance(expected, list) or not expected:
            continue
        mine = view(ours["cases"][index]["region"]["trackedChanges"])
        n = theirs["n"]
        if n in replaced:
            # a tracked Replace: Office JS lists the insertion alone, where the markup holds the deletion
            # too, which is listed here (CR-010's listing)
            assert mine == [*view(expected), ["Deleted", ""]], f"case {n}"
        elif n in (54, 55):
            # Word lists formatting across a table's cells as one change; this package's listing a run at a time
            assert len(expected) == 1, f"case {n}"
            assert len(mine) == 5, f"case {n}"
        elif n == 87:
            # a tracked table insertion at a cell's end: Word's one change is the table and the mark before
            # it; this package lists the mark alone, its range the mark (check 53), and not the nested
            # table's rows
            assert view(expected) == [["Added", "\\u000dNEWCELL\\u000d\\u000a"]]
            assert mine == [["Added", "\\u000d"]]
        else:
            assert mine == view(expected), f"case {n}"
            compared.append(n)
    # case 1's deletion and insertion through a paragraph's mark, each one change spanning two paragraphs, among them
    assert 1 in compared and len(compared) >= 19, compared


def test_cr_014_phase_1_a_range_across_paragraphs_from_expand_to_reads_every_paragraph_the_writes_it_does_not_take_refuse():
    package = loaded("52-ranges.docx")
    span = word_in(package, "HEAD3B").expand_to(word_in(package, "BODY3A"))
    assert span.spans_paragraphs is True
    assert span.text == "HEAD3B\rBODY3A"
    assert [p.text for p in span.paragraphs] == ["HEAD3A HEAD3B", "BODY3A BODY3B"]
    assert span.end_paragraph.text == "BODY3A BODY3B"
    assert [span.get_range(location).text for location in ("Whole", "Content", "Start", "End", "After")] == [
        "HEAD3B\rBODY3A",
        "HEAD3B\rBODY3A",
        "",
        "",
        "",
    ]
    assert span.get_range("End").paragraph.text == "BODY3A BODY3B"
    # the font reads None where the runs differ: the heading's are bold, the body paragraph's italic
    assert isinstance(span.font, MixedFont)
    assert [span.font.bold, span.font.italic, span.font.underline] == [None, None, "None"]
    # search within the span, a paragraph at a time
    assert [r.text for r in span.search("3")] == ["3", "3"]
    assert span.search("HEAD3A") == []
    # refused before anything changes: a hyperlink across paragraphs (Word writes a field across them,
    # core-ts CR-014 section 12), and a control on a range not holding its paragraphs whole, as Word
    # refused it (check 54): Office JS's GeneralException is range.partial_paragraph here (section 4.9 item 1)
    before = package.body.text

    def set_link() -> None:
        span.hyperlink = "https://example.com"

    with pytest.raises(SpanError) as refused:
        set_link()
    assert refused.value.code == "range.spans_paragraphs"
    with pytest.raises(SpanError) as refused:
        span.insert_content_control()
    assert refused.value.code == "range.partial_paragraph"
    assert package.body.text == before
    # get_xml (core-ts's shim's getOoxml) gives its two paragraphs' pieces (to_api_script writes the same blocks)
    ooxml = span.get_xml()
    assert re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", ooxml) == ["HEAD3B", "BODY3A"]
    assert len(re.findall(r"<w:p\b", ooxml)) == 2
    # expand_to from a paragraph's whole range takes its mark; ranges in different stories are refused
    to_mark = word_in(package, "HEAD15B").expand_to(word_in(package, "BODY15A").paragraphs[0].get_range("Whole"))
    assert to_mark.text == "HEAD15B\rBODY15A BODY15B\r"
    assert to_mark.get_range("Content").text == "HEAD15B\rBODY15A BODY15B"
    assert to_mark.get_range("After").paragraph.text == "NEXT15A NEXT15B"
    note = next(p for p in package.main_document_part.footnotes_part.body.paragraphs if "NOTE12" in p.text)
    assert re.search("NOTE12", note.text)
    with pytest.raises(SpanError) as other:
        span.expand_to(note.get_range())
    assert other.value.code == "range.other_story" and "different stories" in str(other.value)


def test_cr_014_phase_1_the_views_across_paragraphs_a_block_control_a_comment_d2_a_tracked_change_a_body_and_a_table_give_one_range():
    package = loaded("52-ranges.docx")
    body = package.body
    # a block control of three paragraphs: one Range
    control = next(c for c in body.content_controls if c.tag == "c52-13")
    assert control.get_range().text == "HEAD13A HEAD13B\rBODY13A BODY13B\rLIST13A LIST13B\r"
    assert control.get_range("Content").text == "HEAD13A HEAD13B\rBODY13A BODY13B\rLIST13A LIST13B"
    assert len(control.get_range().paragraphs) == 3
    # the comment from the body paragraph into the list item (case 1): one Range as Word's, get_ranges()
    # a paragraph at a time
    (comment,) = body.get_comments()
    assert comment.get_range().text == "BODY1B\rLIST1A"
    assert [r.text for r in comment.get_ranges()] == ["BODY1B", "LIST1A"]
    # the tracked insertion through the list item's mark: its range spans both paragraphs, as Word's (case 1)
    inserted = next(c for c in body.get_tracked_changes() if c.type == "Added")
    assert inserted.get_range().text == " LIST1B\rPLAIN1A "
    # a table's whole range ends each row \r\n (case 1), a cell's body's whole range its last paragraph's mark, \t
    table = body.tables[0]
    assert table.get_range().text == "T1R1C1A T1R1C1B\tT1R1C2A T1R1C2B\r\nT1R2C1A T1R2C1B\tT1R2C2A T1R2C2B\r\n"
    assert table.get_range("After").paragraph.text == "END1"
    assert table.cell(0, 0).body.get_range().text == "T1R1C1A T1R1C1B\t"
    # the body's whole range is its text and the last mark
    assert body.get_range().text == f"{body.text}\r"
    assert body.get_range("Content").text == body.text
    # the tables a range holds: one it crosses, none for a range inside a cell
    across = word_in(package, "HEAD8B").expand_to(word_in(package, "BODY8A"))
    assert [t.row_count for t in across.tables] == [2]
    assert table.cell(0, 0).body.get_range().tables == []
    assert [p.table_nesting_level for p in across.paragraphs] == [0, 1, 1, 1, 1, 0]


# ---------------------------------------------------------------------------
# check 53, phase 1's choices (core-ts ef57729)
# ---------------------------------------------------------------------------

#: core-ts's check53Cases() specs (build.mjs at ef57729), in document order.
_CHECK_53_SPECS = [
    "range head{n}b sub{n}a",
    "range body{n}b plain{n}a",
    "range head{n}b head{n}c",
    "range head{n}a head{n}b",
    "range head{n}b head{n}c",
    "para head{n}a",
    "range emph{n}b emph{n}b",
    "range plain{n}a emph{n}b",
    "range plain{n}a plain{n}a",
    "para list{n}a",
    "para plain{n}a",
    "para head{n}a",
    "para list{n}a",
    "para plain{n}a",
    "control c53-{n}",
    "control c53-{n}",
    "control c53-{n}",
    "control c53-{n}",
    "range plain{n}b ctrl{n}a",
    "range ctrl{n}a ctrl{n}b",
    "range plain{n}a plain{n}c",
    "table t{n}r1c1a",
    "cellbody t{n}a",
    "range n{n}a n{n}a",
    *["changes"] * 6,
    "range red{n}a blue{n}a",
    "stories head{n}a",
    "body",
    "placeholder c53-{n}",
    "placeholder c53-{n}",
]


def test_check_53s_file_and_snippet_agree_every_case_parses_to_its_op_every_word_an_op_names_is_in_the_file_once():
    package = load(OFFICE_JS / "53-reads.docx")
    texts = [p.text for p in package.body.paragraphs]
    regions = [r for r in map(checks.parse_region_53, texts) if r is not None]
    assert [
        (r["n"], r["op"] + (" " if r["args"] else "") + " ".join(r["args"]).lower()) for r in regions
    ] == [(i + 1, spec.format(n=i + 1)) for i, spec in enumerate(_CHECK_53_SPECS)]

    def count(token: str) -> int:
        pattern = re.compile(rf"(?<![A-Za-z0-9]){re.escape(token)}(?![A-Za-z0-9])")
        return sum(len(pattern.findall(text)) for text in texts)

    tags = [c.tag for c in package.body.content_controls]
    for r in regions:
        for token in (f"BEGIN{r['n']}", f"END{r['n']}"):
            assert count(token) == 1, f"case {r['n']}: {token}"
        if r["op"] in ("control", "placeholder"):
            assert tags.count(r["args"][0]) == 1, f"case {r['n']}"
        else:
            for token in r["args"]:
                assert count(token) == 1, f"case {r['n']}: {token}"


def test_check_53_word_365_simple_markup_the_snippet_through_the_port_reads_as_word_does_but_for_the_differences_listed_each_outside_cr_014_phase_1():
    ours = checks.run_check_53(loaded("53-reads.docx"), "docx4j-python")
    measured = word("53-reads-word365-simple.json")
    all_markup = word("53-reads-word365-all.json")
    differing: dict[int, list[str]] = {}
    for index, theirs in enumerate(measured["cases"]):
        a, b = flat(ours["cases"][index]["reads"]), flat(theirs["reads"])
        # a colour reads "" for automatic and resolves theme colours another way here (Font's, CR-002 and
        # CR-009): compared below where it is CR-013's
        keys = [k for k in dict.fromkeys([*a, *b]) if k != "font.color" and a.get(k) != b.get(k)]
        if keys:
            differing[theirs["n"]] = sorted(keys)

    def change_read(n: int) -> list[str]:
        return [
            f"paragraphs.{n}.{which}.0.{field}"
            for which in ("changes", "wholeChanges")
            for field in ("type", "text", "range", "rangeParagraphs")
        ]

    placeholder = ["control.text", "control.placeholderText", "control.paragraph", "control.whole", "showingPlaceholderText"]
    style_other = ["paragraph.style", "paragraph.styleBuiltIn", "whole.style", "whole.styleBuiltIn"]
    assert differing == {
        # a single paragraph's font reads its first run (CR-002), where Office JS reads '' or None where
        # its runs differ (core-ts CR-002 section 59, proposed there)
        8: ["font.italic"],
        31: ["font.highlightColor"],
        # a paragraph after a nested table in a cell, whose style isn't Normal: Word's 'Other', not
        # followed (core-ts CR-002 section 57; Q6)
        10: sorted(style_other),
        12: sorted(style_other),
        # tracked changes in Simple Markup. A change's text here is docx4j's, as All Markup reads it,
        # where Simple Markup reads a deletion's as ''. Word joins the paragraphs a deleted mark ends, so
        # its collapsed deletion touches two paragraphs; and its paragraph listing gives each paragraph
        # its part of a change (25), and a deleted paragraph none (30). The listing and the readers are
        # CR-010's and CR-011's (core-ts's CR-002 and CR-012)
        25: sorted(
            [
                "region.0.text",
                "region.0.rangeParagraphs",
                "paragraphs.1.changes.0.text",
                "paragraphs.1.wholeChanges.0.text",
                *change_read(2),
            ]
        ),
        28: sorted(["region.0.text", "paragraphs.1.changes.0.text", "paragraphs.1.wholeChanges.0.text"]),
        30: sorted(["region.0.text", *change_read(2)]),
        # Body.footnotes is not implemented, so the snippet can't reach expand_to's refusal
        32: [""],
        # body.text in Simple Markup joins paragraphs at a deleted mark, drops a deleted paragraph and
        # reads a control showing its placeholder as empty (CR-011's readers, core-ts's CR-012)
        33: sorted(["text.length", "Whole.length", "Content.length"]),
        # a control showing its placeholder reads as empty in Word, and showingPlaceholderText is not
        # implemented (core-ts's too)
        34: sorted(placeholder),
        35: sorted(placeholder),
    }
    assert re.match(r"^error: NotImplemented", ours["cases"][31]["reads"])
    # the colour across two paragraphs whose runs' colours differ: '' (MixedFont), as Word's
    assert flat(ours["cases"][0]["reads"])["font.color"] == flat(measured["cases"][0]["reads"])["font.color"] == ""
    # the changes' text is All Markup's
    for n in (25, 28, 30):
        assert ours["cases"][n - 1]["reads"]["region"][0]["text"] == all_markup["cases"][n - 1]["reads"]["region"][0]["text"], n


# ---------------------------------------------------------------------------
# phase 2, the writes (core-ts 53db5bc, as 87bbb38 rewrote them): CR-013 Phase C
# ---------------------------------------------------------------------------

#: Cases 86 to 95 are core-ts CR-002 section 55's (Word's table rule), which is not implemented here either:
#: phase 2's tests stop at 85.
RANGE_CASES = 85


@functools.cache
def phase_2_run() -> tuple[dict, bytes]:
    """Check 52's snippet in its ``run`` mode, through the port on its own file, once for the tests below."""
    package = loaded("52-ranges.docx")
    result = checks.run_check_52(package, "run")
    return result, package.save()


@functools.cache
def phase_3_run() -> tuple[dict, bytes]:
    """Check 54's snippet in its ``run`` mode, through the port on its own file, once for the tests below."""
    package = loaded("54-writes.docx")
    result = checks.run_check_54(package, "run")
    return result, package.save()


def loaded_bytes(data: bytes):
    package = load(io.BytesIO(data))
    package.get_run_font_selector()
    return package


def _run_view(case: dict, *, links: bool = False) -> dict:
    """What core-ts's run tests compare of a case: under tracking Word's text is All Markup's, so a tracked
    case's paragraphs are compared without their text, and its text, tables and controls not."""
    r = case["region"]
    if isinstance(r, str):
        return {"region": r}
    out: dict[str, Any] = {
        "error": case.get("error"),
        "result": case.get("result"),
        "range": case.get("range"),
        "comments": r.get("comments"),
    }
    if links:
        out["hyperlink"] = r.get("hyperlink")
    # check 52's case 1 is read only, but its content holds tracked changes from the start
    markup = case["tracked"] or (not links and case["n"] == 1)
    paragraphs = r.get("paragraphs")
    out["paragraphs"] = (
        [{**_without_font(p), "text": None} if markup else _without_font(p) for p in paragraphs]
        if isinstance(paragraphs, list)
        else paragraphs
    )
    changes = r.get("trackedChanges")
    out["changes"] = [x["type"] for x in changes] if isinstance(changes, list) else changes
    if not markup:
        out.update(text=r.get("text"), tables=r.get("tables"), controls=r.get("contentControls"))
    return json.loads(json.dumps(out))


def _differing(ours: dict, measured: dict, *, limit: int, links: bool = False) -> dict[int, list[str]]:
    differing: dict[int, list[str]] = {}
    for index, theirs in enumerate(measured["cases"]):
        if theirs["n"] > limit:
            break
        a, b = _run_view(ours["cases"][index], links=links), _run_view(theirs, links=links)
        if isinstance(a.get("paragraphs"), list) and isinstance(b.get("paragraphs"), list):
            _alignment_as_ours(b["paragraphs"], a["paragraphs"], theirs["n"])
        keys = sorted(k for k in set(a) | set(b) if json.dumps(a.get(k), sort_keys=True) != json.dumps(b.get(k), sort_keys=True))
        if keys:
            differing[theirs["n"]] = keys
    return differing


def each(ns: list[int], keys: list[str]) -> dict[int, list[str]]:
    return {n: sorted(keys) for n in ns}


def test_cr_014_phase_2_check_52s_writes_run_through_the_port_on_the_same_file_leave_each_region_as_word_365s_run_did():
    ours, _saved = phase_2_run()
    measured = word("52-ranges-word365-run.json")
    # a paragraph's font reads its first run here (CR-002); Word's 'Left' is this package's "Unknown" where no
    # w:jc says otherwise (CR-002 section 15)
    assert _differing(ours, measured, limit=RANGE_CASES) == {
        # Word's refusals are this package's, under this package's code: range.partial_table for Office JS's
        # GeneralException (CR-013 section 4.9 item 1)
        **each([9, 10, 22, 23, 34, 35, 44, 45], ["error"]),
        # a tracked Replace: Office JS lists the insertion alone, where the markup holds the deletion too, which
        # this package's listing (CR-010) lists; Word's save has it (the markup test below)
        **each([39, 40, 41, 42, 43, 46, 47, 48], ["changes"]),
        # the bold cases' report reads each word through get_text_ranges, which is not implemented (the markup
        # test holds their w:b to Word's); Word lists formatting across a table's cells as one change, this
        # package a run at a time
        **each([49, 50, 51, 52, 53, 56], ["paragraphs"]),
        **each([54, 55], ["paragraphs", "changes"]),
        # on a control's content range Word's insertParagraph adds an empty paragraph as well, which is not
        # copied (core-ts CR-014 section 6.2)
        **each([60, 66], ["paragraphs", "text", "controls"]),
        # insert_comment (core-ts's phase 3): the comments, their ranges and what the snippet reads back are
        # Word's, 79 to 85 among them; core-ts's differ in the snippet's own error, its result asynchronous there
    }
    refused = [c["n"] for c in ours["cases"] if c.get("error") == "error: range.partial_table"]
    assert refused == [9, 10, 22, 23, 34, 35, 44, 45]
    assert all(measured["cases"][n - 1]["error"] == "error: GeneralException" for n in refused)
    # ^p (core-ts's phase 3), case 2's eleven queries: each paragraph's mark outside the table, a hit across two
    # paragraphs' mark, and nothing into or out of the table, at a cell's end, at a row's end, or for ^p^p
    assert ours["cases"][1]["searches"] == measured["cases"][1]["searches"]
    assert len(measured["cases"][1]["searches"]) == 11


def test_cr_014_phase_2_accept_all_after_check_52s_writes_through_the_port_gives_word_365s_accept_all_after_its_run_region_for_region():
    _ours, saved = phase_2_run()
    package = loaded_bytes(saved)
    revisions.accept_all(package)
    back = loaded_bytes(package.save())
    ours = checks.run_check_52(back, "report")
    measured = word("52-ranges-word365-accepted.json")
    differing = []
    for index, theirs in enumerate(measured["cases"]):
        if theirs["n"] > RANGE_CASES:
            break

        # Range.footnotes is not implemented, and a paragraph's font reads its first run here (CR-002)
        def view(region: dict) -> dict:
            region = {k: v for k, v in region.items() if k != "footnotes"}
            if isinstance(region.get("paragraphs"), list):
                region["paragraphs"] = [_without_font(p) for p in region["paragraphs"]]
            return json.loads(json.dumps(region))

        a, b = view(ours["cases"][index]["region"]), view(theirs["region"])
        if isinstance(a.get("paragraphs"), list) and isinstance(b.get("paragraphs"), list):
            _alignment_as_ours(b["paragraphs"], a["paragraphs"], theirs["n"])
        if json.dumps(a, sort_keys=True) != json.dumps(b, sort_keys=True):
            differing.append(theirs["n"])
    # every tracked deletion and replacement, accepted, is Word's, and every comment across paragraphs (phase 3);
    # what differs is listed in the test above: get_text_ranges (the bold cases), Word's extra empty paragraph in a
    # control
    assert differing == [49, 50, 51, 52, 53, 54, 55, 56, 60, 66]


_IGNORED_KEYS = frozenset({"rsid_r", "rsid_r_pr", "rsid_r_default", "rsid_p", "rsid_del", "para_id", "text_id", "b_cs", "i_cs", "sz_cs", "lang"})


def _camel(name: str) -> str:
    """A model field's name as core-ts's markup lines name it (``r_fonts`` ``rFonts``, ``del_value`` ``del``)."""
    if name == "del_value":
        return "del"
    head, *rest = name.split("_")
    return head + "".join(word[:1].upper() + word[1:] for word in rest)


def _keys(value: Any) -> str:
    if value is None:
        return ""
    names = sorted(
        _camel(field.name)
        for field in dataclasses.fields(value)
        if field.name not in _IGNORED_KEYS and field.name != "r_pr_change" and getattr(value, field.name) is not None
    )
    return ",".join(names) + ("+change" if getattr(value, "r_pr_change", None) is not None else "")


def _local(element: Any) -> str:
    return (element_name(element) or "").rsplit("}", 1)[-1]


def _markup_runs(items: Any, revision: str = "") -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for item in items or ():
        name = _local(item)
        if name in ("ins", "del"):
            out += _markup_runs(item.content, name)
            continue
        if name == "hyperlink":
            out += _markup_runs(item.content, f"{revision}link")
            continue
        if name != "r":
            if name.startswith(("comment", "bookmark")):
                out.append({"revision": name, "rPr": "", "text": ""})
            continue
        text = ""
        for child in item.content or ():
            local = _local(child)
            if local in ("t", "delText"):
                text += getattr(child, "value", "") or ""
            elif local != "lastRenderedPageBreak":
                text += f"<{local}>"
        out.append({"revision": revision, "rPr": _keys(item.r_pr), "text": text})
    merged: list[dict[str, str]] = []
    for run in out:
        if merged and merged[-1]["revision"] == run["revision"] and merged[-1]["rPr"] == run["rPr"]:
            merged[-1]["text"] += run["text"]
        else:
            merged.append(dict(run))
    return merged


def _val(value: Any) -> Any:
    found = getattr(value, "val", None)
    return getattr(found, "value", found)


def _markup_lines(element: Any, indent: str = "") -> list[str]:
    name = _local(element)
    if name == "p":
        p_pr = element.p_pr
        props: list[str] = []
        if p_pr is not None:
            for value in (
                _val(p_pr.p_style) if p_pr.p_style is not None else None,
                _val(p_pr.jc) if p_pr.jc is not None else None,
                "num" if p_pr.num_pr is not None else "",
                "pPrChange" if p_pr.p_pr_change is not None else "",
            ):
                if value:
                    props.append(str(value))
        content = " ".join(
            f"{run['revision'] + ':' if run['revision'] else ''}{json.dumps(run['text'])}"
            f"{'{' + run['rPr'] + '}' if run['rPr'] else ''}"
            for run in _markup_runs(element.content)
        )
        return [f"{indent}P[{'/'.join(props)}|{_keys(p_pr.r_pr if p_pr is not None else None)}] {content}"]
    if name == "tbl":
        out: list[str] = []
        for tr in element.content or ():
            if _local(tr) != "tr":
                continue
            out.append(f"{indent}TR {_keys(tr.tr_pr)}")
            for tc in tr.content or ():
                for block in getattr(tc, "content", None) or ():
                    out += _markup_lines(block, indent + "  ")
        return out
    if name == "sdt":
        out = [f"{indent}SDT"]
        for block in element.sdt_content.content or ():
            out += _markup_lines(block, indent + "  ")
        return out
    return [f"{indent}<{name}>"]


def markup_of(package) -> dict[int, list[str]]:
    """A region's markup, a line a paragraph, as core-ts's ``markupOf`` writes it (``ranges.test.mjs`` at ``87bbb38``).

    Its style, alignment and numbering, its mark's properties (``ins``, ``del``, ``b``, ``+change`` for its
    w:rPrChange) and a w:pPrChange, then its runs, those with the same revision and properties merged, each
    ``ins:`` or ``del:`` with its properties; a table's rows (``TR del``) and cells' paragraphs. Ids, dates, rsids,
    paraIds, w:lastRenderedPageBreak and the complex-script twins (bCs, iCs, szCs) are left out. The model's
    field names are core-ts's (``rFonts``, ``del``), so that its pinned lines compare.
    """

    def text_of(element: Any) -> str:
        return "".join(run["text"] for run in _markup_runs(element.content)) if _local(element) == "p" else ""

    regions: dict[int, list[str]] = {}
    n = None
    for element in package.main_document_part.contents.body.content:
        begin = re.match(r"^BEGIN(\d+) ", text_of(element))
        if begin:
            n = int(begin.group(1))
            regions[n] = []
            continue
        if n is not None and text_of(element) == f"END{n}":
            n = None
            continue
        if n is not None:
            regions[n] += _markup_lines(element)
    return regions


def test_cr_014_phase_2_the_markup_check_52s_writes_leave_through_the_port_is_word_365s_in_its_save_after_the_run_region_for_region():
    mine = markup_of(loaded_bytes(phase_2_run()[1]))
    measured = markup_of(load(OFFICE_JS / "52-ranges-word365.docx"))
    differing = [n for n in measured if n <= RANGE_CASES and mine.get(n) != measured.get(n)]
    # 1, read only: Word's save gave the comment reference's run w:u and w:sz of its own; 13: Word's filler for
    # an emptied control, a bold run of five spaces (core-ts CR-014 section 9.4); 60 and 66: Word's extra empty
    # paragraph in a control; 67 and 68: Word gives a paragraph's mark its last run's formatting (italic) when
    # insert_paragraph("After") goes after it, not followed (core-ts CR-014 section 10); 79 to 85: insert_comment
    # (phase 3), whose markers are Word's but for the two differences below
    assert differing == [1, 13, 60, 66, 67, 68, 79, 80, 81, 82, 83, 84, 85]
    # the comments across paragraphs (phase 3): the start marker before the range's first character, the end and
    # the reference run after its last, in its last paragraph, as Word's. Word gives the reference run the text's
    # size and, where text follows it, that text's italic, over the CommentReference style; this package's
    # reference run has the style alone, as within one paragraph. And on a control's content range (82) Word put
    # the start before the control's first paragraph, which reads the same (core-ts CR-014 section 9.2)
    for n in (79, 80, 81, 82, 83, 84, 85):
        expected = [
            re.sub(r'"<commentReference>"\{(i,)?rStyle,sz\}', '"<commentReference>"{rStyle}', line) for line in measured[n]
        ]
        if n == 82:
            expected = [
                line.replace('  P[Heading1|] "HEAD82A', '  P[Heading1|] commentRangeStart:"" "HEAD82A')
                for line in expected
                if line != "  <commentRangeStart>"
            ]
        assert mine[n] == expected, f"case {n}"
    assert measured[13] == ["SDT", '  P[ListParagraph/num|] "     "{b}']
    assert mine[13] == ["SDT", "  P[ListParagraph/num|] "]
    for n in (67, 68):
        assert [re.sub(r'^P\[center\|\] "BODY', 'P[center|i] "BODY', line) for line in mine[n]] == measured[n], f"case {n}"
    # what Word wrote, which this package now writes: a tracked deletion across paragraphs marks every mark but the
    # last deleted, with no w:pPrChange (case 16); a tracked Replace puts its w:ins before the w:del (case 39); an
    # untracked one gives its text the heading's look in a body paragraph (case 29); bold skips the heading's runs
    # and bolds a mark in the range, tracked with its w:rPrChange (case 56); a paragraph split tracked (case 70)
    assert measured[16] == mine[16] == ['P[Heading1|del] "HEAD16A " del:"HEAD16B"', 'P[center|] del:"BODY16A"{i} " BODY16B"{i}']
    assert measured[39] == mine[39] == [
        'P[Heading1|del] "HEAD39A " ins:"NEWTEXT" del:"HEAD39B"',
        'P[center|] del:"BODY39A"{i} " BODY39B"{i}',
    ]
    assert measured[29] == mine[29] == ['P[center|] "HEAD29A " "NEWTEXT"{b,color,rFonts,sz} " BODY29B"{i}']
    assert measured[56][1:3] == mine[56][1:3] == ['  P[Heading1|] "HEAD56A HEAD56B"', '  P[center|b+change] "BODY56A BODY56B"{b,i+change}']
    assert measured[70][1:] == mine[70][1:] == [
        'P[center|i,ins] "BODY70A"{i}',
        'P[center|i,ins] ins:"NEWPARA"{i}',
        'P[center|] " BODY70B"{i}',
    ]


def _region(package, n: int) -> list:
    paragraphs = package.body.paragraphs
    begin = next(i for i, p in enumerate(paragraphs) if p.text.startswith(f"BEGIN{n} "))
    end = next(i for i, p in enumerate(paragraphs) if i > begin and p.text == f"END{n}")
    return paragraphs[begin + 1 : end]


def test_cr_014_phase_2_delete_and_replace_across_paragraphs_and_the_range_a_write_leaves_reject_all_after_a_tracked_write_gives_back_what_was_there():
    package = loaded("52-ranges.docx")

    def texts(n: int) -> list[str]:
        return [p.text for p in _region(package, n)]

    # untracked: the paragraphs join in the last one's properties, and the range collapses where they joined
    span = word_in(package, "HEAD3B").expand_to(word_in(package, "BODY3A"))
    span.delete()
    assert texts(3) == ["HEAD3A  BODY3B"]
    assert [_region(package, 3)[0].alignment, _region(package, 3)[0].style_built_in] == ["Centered", "Normal"]
    assert [span.spans_paragraphs, span.text, span.start] == [False, "", 7]
    # Replace: the range covers the new text; "" is a deletion that keeps a last mark
    replaced = word_in(package, "HEAD38B").expand_to(word_in(package, "BODY38A").paragraphs[0].get_range("Whole"))
    inserted = replaced.insert_text("NEW", location="Replace")
    assert [inserted.text, replaced.text, texts(38)] == ["NEW", "NEW", ["HEAD38A NEW", "NEXT38A NEXT38B"]]
    word_in(package, "HEAD29B").expand_to(word_in(package, "BODY29A")).insert_text("", location="Replace")
    assert texts(29) == ["HEAD29A  BODY29B"]
    # a footnote whose reference a deletion takes leaves the notes part (check 52 case 12): when the write ends
    # here, and so on save (core-ts queues it for the save)

    def notes_of(p) -> int:
        notes = p.main_document_part.footnotes_part.contents.footnote
        return sum(1 for f in notes if getattr(getattr(f, "type_value", None), "value", None) in (None, "normal"))

    notes = notes_of(package)
    word_in(package, "HEAD12B").expand_to(word_in(package, "BODY12A")).delete()
    assert notes_of(package) == notes - 1
    assert notes_of(load(io.BytesIO(package.save()))) == notes - 1

    # tracked: one deletion across the paragraphs; Reject All gives back the paragraphs, Accept All the untracked join
    original = texts(16)
    package.change_tracking_mode = "TrackAll"
    tracked = word_in(package, "HEAD16B").expand_to(word_in(package, "BODY16A"))
    tracked.delete()
    assert [(c.type, c.text) for c in package.body.get_tracked_changes() if "16" in c.text] == [("Deleted", "HEAD16B\rBODY16A")]
    assert [tracked.spans_paragraphs, tracked.text] == [False, ""], "collapsed where the deletion starts, as Simple Markup reads it"
    word_in(package, "HEAD49B").expand_to(word_in(package, "BODY49A")).font.bold = True
    word_in(package, "HEAD57B").expand_to(word_in(package, "BODY57A")).insert_paragraph("NEW", location="Before")
    word_in(package, "HEAD39B").expand_to(word_in(package, "BODY39A")).insert_text("NEW", location="Replace")
    package.change_tracking_mode = "Off"
    assert texts(57) == ["HEAD57A ", "NEW", "HEAD57B", "BODY57A BODY57B"]
    rejected = loaded_bytes(package.save())
    revisions.reject_all(rejected)
    assert [p.text for p in _region(rejected, 16)] == original
    assert [(p.text, p.style_built_in) for p in _region(rejected, 57)] == [("HEAD57A HEAD57B", "Heading1"), ("BODY57A BODY57B", "Normal")]
    assert [p.text for p in _region(rejected, 39)] == ["HEAD39A HEAD39B", "BODY39A BODY39B"]
    # the bold rejected: no run and no mark keeps a w:b or a change
    xml = rejected.main_document_part.xml.decode()
    region_49 = xml[xml.index("BEGIN49 ") : xml.index("END49")]
    assert not re.search(r"<w:b/>|w:rPrChange", region_49)


def test_cr_014_phase_2_whole_tables_a_cell_the_body_and_a_control_what_word_refuses_is_refused_before_anything_changes():
    package = loaded("52-ranges.docx")
    body = package.body
    # a range starting or ending inside a table: Word's refusal (GeneralException there, range.partial_table here),
    # tracked or not, and nothing changed
    before = body.text
    for mode in ("Off", "TrackAll"):
        package.change_tracking_mode = mode
        for start, end in (("HEAD9B", "T9R2C1A"), ("T10R1C1B", "BODY10A")):
            span = word_in(package, start).expand_to(word_in(package, end))
            with pytest.raises(SpanError) as refused:
                span.delete()
            assert refused.value.code == "range.partial_table" and "inside a table" in str(refused.value)
            with pytest.raises(SpanError) as refused:
                span.insert_text("X", location="Replace")
            assert refused.value.code == "range.partial_table"
    package.change_tracking_mode = "Off"
    assert body.text == before
    # Word applies a font and a paragraph to such a range (check 52 cases 51 and 59)
    cell = word_in(package, "T51R1C1B").expand_to(word_in(package, "BODY51A"))
    cell.font.bold = True
    assert cell.font.bold is True
    # a range holding a whole block control: under tracking the control is deleted with Word's markers (check 54
    # case 15, core-ts CR-014 section 12), Accept All taking it and joining the paragraphs, Reject All giving them
    # back; untracked, the control goes as a table does and the paragraphs either side join
    tracked = loaded_bytes(package.save())
    tracked.change_tracking_mode = "TrackAll"
    word_in(tracked, "BEGIN26").expand_to(word_in(tracked, "END26")).delete()
    tracked.change_tracking_mode = "Off"
    assert re.search(
        r"<w:customXmlDelRangeStart[^>]*/><w:sdt>.*?<w:sdtContent><w:customXmlDelRangeEnd", tracked.main_document_part.xml.decode()
    )
    saved = tracked.save()
    accepted, rejected = loaded_bytes(saved), loaded_bytes(saved)
    revisions.accept_all(accepted)
    revisions.reject_all(rejected)
    assert [c for c in accepted.body.content_controls if c.tag == "c52-26"] == []
    assert [p for p in accepted.body.paragraphs if re.search("BEGIN26|END26", p.text)] == []

    # the region as it was (Reject All rejects case 1's changes too, so the region alone is compared)
    def region_26(document) -> str:
        text = document.body.text
        return text[text.index("BEGIN26 ") : text.index("END26")]

    assert region_26(rejected) == region_26(package)
    around = word_in(package, "BEGIN26").expand_to(word_in(package, "END26"))
    controls, paragraphs = len(body.content_controls), len(body.paragraphs)
    around.delete()
    # BEGIN26, the control's three paragraphs and END26 become one empty paragraph
    assert [len(body.content_controls), len(body.paragraphs)] == [controls - 1, paragraphs - 4]
    assert [p for p in body.paragraphs if re.search("BEGIN26|END26", p.text)] == []
    # into a block content control from outside it: the texts go and nothing joins across the control's edge; the
    # start's paragraph takes the end's properties (check 54 case 18)
    control = next(c for c in body.content_controls if c.tag == "c52-13")
    begin = word_in(package, "BEGIN13").paragraphs[0]
    word_in(package, "BEGIN13").expand_to(control.paragraphs[0].search("HEAD13A")[0]).delete()
    assert [begin.text, begin.style_built_in] == ["", "Heading1"]
    assert [control.paragraphs[0].text, control.paragraphs[0].style_built_in] == [" HEAD13B", "Heading1"]
    # a table's whole range takes the table; a cell's whole range keeps its last paragraph, emptied (CR-002 section 55)
    tables = len(body.tables)
    body.tables[0].get_range().delete()
    assert len(body.tables) == tables - 1
    cell_body = body.tables[0].cell(0, 0).body
    cell_body.get_range().delete()
    assert [p.text for p in cell_body.paragraphs] == [""]
    # the control's content: one paragraph left, in its last paragraph's properties (check 52 case 13)
    control.get_range("Content").delete()
    assert [(p.text, p.style_built_in) for p in control.paragraphs] == [("", "ListParagraph")]
    # the body's content: one empty paragraph, in the last one's properties
    body.get_range("Content").delete()
    assert [p.text for p in body.paragraphs] == [""]
    assert body.tables == []


def test_cr_014_phase_2_insert_paragraph_splits_a_paragraph_inside_a_link_into_two_links_and_refuses_a_split_inside_a_content_control_insert_text_at_the_ends_replace_text_a_paragraph_at_a_time():
    package = create_package()
    body = package.body
    body.insert_xml(
        f'<w:p xmlns:w="{W}"><w:pPr><w:jc w:val="center"/></w:pPr>'
        '<w:hyperlink w:anchor="here"><w:r><w:t>linked text</w:t></w:r></w:hyperlink><w:r><w:t> after</w:t></w:r></w:p>',
        location="End",
    )
    body.insert_xml(
        f'<w:p xmlns:w="{W}"><w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
        "<w:sdtContent><w:r><w:t>in a control</w:t></w:r></w:sdtContent></w:sdt></w:p>",
        location="End",
    )
    body.insert_paragraph("last one", location="End")
    linked, controlled = [p for p in body.paragraphs if p.text][:2]
    span = linked.search("text")[0].expand_to(body.search("last")[0])
    span.insert_paragraph("NEW", location="Before")
    assert [(p.text, p.alignment) for p in body.paragraphs if p.text][:3] == [
        ("linked ", "Centered"),
        ("NEW", "Centered"),
        ("text after", "Centered"),
    ]
    assert [p.get_range().hyperlink for p in body.paragraphs if p.text][:3] == ["#here", "", "#here"]
    assert [span.text, span.start] == ["text after\rin a control\rlast", 0]
    # a split inside an inline control would make two of it: refused, nothing changed
    across = controlled.search("control")[0].expand_to(body.search("last")[0])
    xml = package.main_document_part.xml.decode()
    with pytest.raises(SpanError) as refused:
        across.insert_paragraph("X", location="Before")
    assert refused.value.code == "range.splits_control"
    assert package.main_document_part.xml.decode() == xml
    # insert_text at the start and the end; replace_text in each paragraph
    span.insert_text("<", location="Start")
    span.insert_text(">", location="End")
    assert span.text == "text after\rin a control\rlast"
    assert "|".join(t for t in body.text.split("\r") if t) == "linked |NEW|<text after|in a control|last> one"
    assert span.replace_text("t", "T") == 5
    assert span.text == "TexT afTer\rin a conTrol\rlasT"


# ---------------------------------------------------------------------------
# check 54 (core-ts 87bbb38): phase 2's choices, phase 3's controls
# ---------------------------------------------------------------------------

#: check 54's cases, as core-ts's build.mjs (87bbb38) lists them: (kind, op, tracking).
_CASES_54 = [
    ("IN1", "paragraph-before", "untracked"), ("IN1", "paragraph-after", "untracked"), ("IN1", "paragraph-before", "tracked"),
    ("IN1", "replace", "tracked"),
    ("TE", "delete", "untracked"), ("TE", "delete", "tracked"), ("TE", "replace", "untracked"),
    ("WT", "delete", "untracked"), ("WT", "delete", "tracked"), ("WL", "delete", "untracked"), ("WL", "delete", "tracked"),
    ("WC", "delete", "untracked"), ("WC", "delete", "tracked"),
    ("CCW", "delete", "untracked"), ("CCW", "delete", "tracked"), ("CCW", "replace", "untracked"), ("CCW", "replace", "tracked"),
    ("CCI", "delete", "untracked"), ("CCI", "delete", "tracked"), ("CCO", "delete", "untracked"),
    ("QP", "italic", "untracked"), ("QP", "italic", "tracked"), ("HP", "size14", "untracked"), ("HP", "size14", "tracked"),
    ("LP", "underline", "untracked"),
    ("PS", "replace", "untracked"),
    ("CMT", "delete", "untracked"), ("CMT", "delete", "tracked"), ("CMS", "delete", "untracked"), ("BMK", "delete", "untracked"),
    ("OWN", "own-delete", "tracked"),
    ("BL", "style", "untracked"), ("BL", "style", "tracked"), ("HB", "hyperlink", "untracked"), ("HB", "ooxml", "untracked"),
    ("HB", "control", "untracked"), ("W2", "control", "untracked"), ("WP", "control", "untracked"), ("TBL", "control", "untracked"),
    ("INTO", "control", "untracked"), ("CELL", "control", "untracked"), ("HB", "control", "tracked"), ("WP", "control", "tracked"),
]  # fmt: skip

#: check 54's ranges, as build.mjs's KINDS write them for case n.
_RANGES_54 = {
    "IN1": "word head{n}b", "TE": "from head{n}b to-table t{n}r1c1a", "WT": "whole body{n}a", "WL": "whole list{n}a",
    "WC": "whole list{n}a", "CCW": "from head{n}b to body{n}a", "CCI": "from head{n}b to body{n}a",
    "CCO": "from body{n}b to plain{n}a", "QP": "from quote{n}b to plain{n}a", "HP": "from head{n}b to plain{n}a",
    "LP": "from list{n}b to plain{n}a", "PS": "from plain{n}b to styl{n}a", "CMT": "from head{n}a to body{n}a",
    "CMS": "from head{n}a to body{n}a", "BMK": "from head{n}a to body{n}a", "OWN": "from head{n}b to own{n}a",
    "BL": "from body{n}b to list{n}a", "HB": "from head{n}b to body{n}a", "W2": "from head{n}b to-mark body{n}a",
    "WP": "start head{n}a to-mark body{n}a", "TBL": "from head{n}b to body{n}a", "INTO": "from head{n}b to t{n}r2c1a",
    "CELL": "from t{n}r1c1b to body{n}a",
}  # fmt: skip


def check_54_cases() -> list[tuple[int, str, str, bool, str]]:
    """``check54Cases()``: (n, kind, op, tracked, spec), in document order."""
    return [
        (n, kind, op, tracking == "tracked", _RANGES_54[kind].format(n=n))
        for n, (kind, op, tracking) in enumerate(_CASES_54, start=1)
    ]


def test_check_54s_file_and_snippet_agree_every_case_parses_to_the_case_it_was_built_as_every_word_its_range_names_is_in_the_file_once():
    package = loaded("54-writes.docx")
    texts = [p.text for p in package.body.paragraphs]
    regions = [region for text in texts if (region := checks.parse_region_54(text)) is not None]
    cases = check_54_cases()
    assert len(cases) == 43
    assert [(r["n"], r["kind"], r["op"], r["tracked"], r["spec"]) for r in regions] == cases

    def count(token: str) -> int:
        return sum(len(re.findall(rf"(?<![A-Za-z0-9]){token}(?![A-Za-z0-9])", text)) for text in texts)

    for r in regions:
        # the OWN case's last word is in the paragraph its run inserts first
        tokens = [f"BEGIN{r['n']}", f"END{r['n']}", *r["range"].values()]
        for token in tokens:
            if not token.startswith("OWN"):
                assert count(token) == 1, f"case {r['n']}: {token}"
    # the snippet runs every case through the port with no error of its own: what fails is this package's
    # refusals, under its codes (Office JS's GeneralException is range.partial_paragraph, range.replaces_table;
    # a hyperlink across paragraphs range.spans_paragraphs), and getBookmarks, which this package doesn't have
    ours = checks.run_check_54(loaded("54-writes.docx"), "run")
    assert len(ours["cases"]) == 43
    errors = {
        value
        for c in ours["cases"]
        for value in [c.get("error"), *(c["region"].values() if isinstance(c["region"], dict) else [c["region"]])]
        if isinstance(value, str) and value.startswith("error:")
    }
    assert errors == {
        "error: range.partial_paragraph",
        "error: range.replaces_table",
        "error: range.spans_paragraphs",
        "error: NotImplemented: Range.getBookmarks",
    }


def test_cr_014_phase_3_a_comment_across_paragraphs_its_range_the_range_commented_p_finds_paragraph_marks_as_word_does():
    package = loaded("52-ranges.docx")
    body = package.body
    # a comment from a heading's middle across a table into a body paragraph: one Range, as Word's (check 52 case 80)
    span = word_in(package, "HEAD80B").expand_to(word_in(package, "BODY80A"))
    comment = span.insert_comment("NEW")
    assert comment.get_range().text == span.text
    assert [c.content for c in span.get_comments()] == ["NEW"]
    # a range starting at a paragraph's end and one ending at the next paragraph's start: the markers stand there
    at_end = word_in(package, "HEAD3B").paragraphs[0].get_range("End").expand_to(word_in(package, "BODY3A"))
    assert at_end.insert_comment("END").get_range().text == "\rBODY3A"
    to_start = word_in(package, "HEAD16B").expand_to(word_in(package, "BODY16A").paragraphs[0].get_range("Start"))
    assert to_start.insert_comment("START").get_range().text == "HEAD16B\r"
    saved = loaded_bytes(package.save())
    assert [c.get_range().text for c in saved.body.get_comments() if c.content in ("NEW", "END", "START")] == [
        "\rBODY3A",
        "HEAD16B\r",
        span.text,
    ]

    # ^p: each paragraph's own mark the range holds, a hit across a mark, nothing at a cell's end or across a table's edge
    region = word_in(package, "BEGIN2").paragraphs[0].get_range("Whole").expand_to(word_in(package, "END2").paragraphs[0].get_range("Whole"))
    assert len(region.search("^p")) == 6
    assert [(r.text, len(r.paragraphs)) for r in region.search("HEAD2B^pBODY2A")] == [("HEAD2B\rBODY2A", 2)]
    assert [len(region.search(q)) for q in ("LIST2B^pT2R1C1A", "T2R1C1B^p", "T2R2C2B^pPLAIN2A", "^p^p")] == [0, 0, 0, 0]
    hit = region.search("^p")[1]
    assert [hit.text, hit.paragraph.text, hit.holds_end_mark] == ["\r", "HEAD2A HEAD2B", True]
    # a paragraph's search reads its own mark; the body's every mark outside a table
    assert [r.text for r in word_in(package, "HEAD2A").paragraphs[0].search("^p")] == ["\r"]
    assert len(body.search("B^pT", match_case=True)) == 0, "into a table from any paragraph before one"
    assert len(body.search("B^pE", match_case=True)) > 80, "into each END paragraph"
    assert len(body.search("HEAD2B^pBODY2A")) == 1
    # a cell's paragraphs: a mark between two of them is a paragraph's; the cell's last isn't
    cell = body.tables[0].cell(0, 0).body
    cell.paragraphs[0].insert_paragraph("SECOND", location="After")
    assert [r.paragraph.text for r in cell.get_range().search("^p")] == ["T1R1C1A T1R1C1B"]
    assert [r.text for r in cell.get_range().search("T1R1C1B^pSECOND")] == ["T1R1C1B\rSECOND"]
    # an empty paragraph's mark after another mark: ^p^p
    word_in(package, "HEAD2B").paragraphs[0].insert_paragraph("", location="After")
    assert [r.text for r in region.search("^p^p")] == ["\r\r"]
    assert [r.text for r in region.search("HEAD2B^p^pBODY2A")] == ["HEAD2B\r\rBODY2A"]
    # wildcards don't read ^p as a mark, but as the two characters: the twelve in case 2's BEGIN paragraph, which
    # lists the queries; a replacement whose query holds one is refused, unmeasured (range.replaces_mark here)
    assert [r.text for r in region.search("^p", match_wildcards=True)] == ["^p"] * 12
    for replace in (
        lambda: body.replace_text("^p", " "),
        lambda: region.replace_text("^p", " "),
        lambda: region.paragraphs[0].replace_text("^p", " "),
    ):
        with pytest.raises(SpanError) as refused:
            replace()
        assert refused.value.code == "range.replaces_mark" and "^p" in str(refused.value)


def test_cr_014_phase_3_a_comments_markers_go_beside_runs_holding_nothing_outside_its_range_a_tab_sharing_a_run_is_split_off_within_one_paragraph_and_across_two():
    package = create_package()
    body = package.body
    xml = f'<w:p xmlns:w="{W}"><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:tab/><w:t>c</w:t></w:r></w:p>'
    body.insert_xml(xml + xml, location="End")
    first, second = [p for p in body.paragraphs if p.text]
    # within one paragraph: from 'b' to its end, where the run starts at 'a' (CR-013 Phase C: the defect was
    # here too, and the marker went beside the whole run, the comment's range "a\tb\tc")
    one = Range(first, 2, 5)
    assert one.insert_comment("ONE").get_range().text == "b\tc"
    # across two: from the first paragraph's 'c' to the second's 'a'
    across = Range(first, 4, 4).expand_to(second.search("a")[0])
    assert across.text == "c\ra"
    assert across.insert_comment("TWO").get_range().text == "c\ra"
    assert [first.text, second.text] == ["a\tb\tc", "a\tb\tc"]


def test_check_54s_run_and_its_report_after_accept_all_are_of_this_file_and_its_saves_are_words_as_copied():
    cases = [(n, kind, op, tracked) for n, kind, op, tracked, _spec in check_54_cases()]
    for name, mode in (("54-writes-word365-run.json", "run"), ("54-writes-word365-accepted.json", "report")):
        measured = word(name)
        assert measured["mode"] == mode
        assert [(c["n"], c["kind"], c["op"], c["tracked"]) for c in measured["cases"]] == cases
    # core-ts 87bbb38's test/check54/manifest.json
    expected = {
        "54-writes-word365.docx": "b5e958f9204751488625ca5adafea3d605a89933",
        "54-writes-word365-accepted.docx": "5047a128aec36d2933d044ca6ae9b7ce62510daf",
    }
    for name, digest in expected.items():
        assert hashlib.sha1((OFFICE_JS / name).read_bytes()).hexdigest() == digest, name


def test_cr_014_phase_3_check_54s_writes_run_through_the_port_on_the_same_file_leave_each_region_as_word_365s_run_did():
    ours, _saved = phase_3_run()
    measured = word("54-writes-word365-run.json")
    assert _differing(ours, measured, limit=99, links=True) == {
        # a tracked Replace within one paragraph: Office JS lists the insertion alone, this package the deletion too
        4: ["changes"],
        # Word's refusals, under this package's codes (CR-013 section 4.9 item 1): replacing a range that ends with
        # a whole table (range.replaces_table), a control on a range not holding its paragraphs whole
        # (range.partial_paragraph)
        **each([7, 36, 37, 39, 40, 41, 42], ["error"]),
        # an emptied list item in a cell reads Word's style 'Other' (core-ts CR-002 section 57's family; Q6)
        **each([10, 11], ["paragraphs"]),
        # Word lists a control's last paragraph's text and mark as two deletions, this package one (CR-010's listing)
        13: ["changes"],
        # a tracked Replace over a whole control: Office JS lists the insertion alone, this package the deletion too
        17: ["changes"],
        # tracked, Word's comment range reads the deleted word (All Markup)
        28: ["comments"],
        # Word's model marks the first paragraph's mark inserted when it inserts one after it, so deleting into the
        # new paragraph took its own mark back and joined them; this package marks the new one's (CR-003 section
        # 4): the same document once accepted or rejected (the Accept All test below)
        31: ["paragraphs"],
        # a style across paragraphs: Word also took away direct formatting (centring, numbering), not followed, and
        # listed one change
        32: ["paragraphs"],
        33: ["changes", "paragraphs"],
        # a hyperlink across paragraphs: Word wrote a HYPERLINK field across them, refused here
        34: ["error", "hyperlink", "paragraphs", "result"],
        # insert_ooxml returns the paragraph the runs went into (CR-003 section 3.2), where Office JS returns the
        # range of what it inserted; core-ts's differs here too, its result asynchronous
        35: ["result"],
        # a tracked control around paragraphs: Word lists its two insertion markers as changes, this package nothing
        43: ["changes"],
    }
    assert [measured["cases"][n - 1]["error"] for n in (7, 36, 37, 39, 40, 41, 42)] == ["error: GeneralException"] * 7
    assert ours["cases"][6]["error"] == "error: range.replaces_table"
    assert {ours["cases"][n - 1]["error"] for n in (36, 37, 39, 40, 41, 42)} == {"error: range.partial_paragraph"}


def test_cr_014_phase_3_accept_all_after_check_54s_writes_through_the_port_is_word_365s_accept_all_in_each_regions_report_and_markup():
    _ours, saved = phase_3_run()
    package = loaded_bytes(saved)
    revisions.accept_all(package)
    back = loaded_bytes(package.save())
    ours = checks.run_check_54(back, "report")
    measured = word("54-writes-word365-accepted.json")
    reported = []
    for index, theirs in enumerate(measured["cases"]):
        # getBookmarks is not implemented
        def view(region: dict) -> dict:
            return json.loads(json.dumps({k: v for k, v in region.items() if k != "bookmarks"}))

        a, b = view(ours["cases"][index]["region"]), view(theirs["region"])
        if isinstance(a.get("paragraphs"), list) and isinstance(b.get("paragraphs"), list):
            _alignment_as_ours(b["paragraphs"], a["paragraphs"], theirs["n"])
        if json.dumps(a, sort_keys=True) != json.dumps(b, sort_keys=True):
            reported.append(theirs["n"])
    mine = markup_of(back)
    words = markup_of(load(OFFICE_JS / "54-writes-word365-accepted.docx"))
    marked = [n for n in words if mine.get(n) != words.get(n)]
    # every case's result is Word's, untracked or accepted, the tracked deletions of a whole control among them (15,
    # 17), but for: a style's resets (32, 33) and a hyperlink across paragraphs (34), not followed; the reference
    # run's size (29); Word's 'Other' for an emptied list item in a cell (10, 11, its report alone); and three rules
    # of Word's Accept All that this package's, docx4j's, doesn't have (9, 13, 19; CR-013 section 13.3)
    assert reported == [9, 10, 11, 13, 19, 32, 33, 34]
    assert marked == [9, 13, 19, 29, 32, 33, 34]
    # Word's Accept All: an emptied paragraph whose deleted mark had a table after it goes (9); a control's emptied
    # last paragraph keeps a bare mark (13); a paragraph whose deleted mark had a control after it takes the
    # control's first paragraph's properties, its content not joined (19). This package keeps the first, keeps the
    # second's properties, and the third's own, as docx4j's resolution does
    assert words[9][0] == "TR "
    assert mine[9][0] == "P[center|] "
    assert words[13][3] == "  P[|] "
    assert mine[13][3] == "  P[ListParagraph/num|] "
    assert words[19][0] == 'P[center|] "HEAD19A "'
    assert mine[19][0] == 'P[Heading1|] "HEAD19A "'


def test_cr_014_phase_3_the_markup_check_54s_writes_leave_through_the_port_is_word_365s_in_its_save_after_the_run_region_for_region():
    mine = markup_of(loaded_bytes(phase_3_run()[1]))
    measured = markup_of(load(OFFICE_JS / "54-writes-word365.docx"))
    differing = [n for n in measured if mine.get(n) != measured.get(n)]
    # 28: the reference run's size, which Word's save gave the file's own reference run (as check 52's case 1); the
    # run itself is in the w:del, as Word's (core-ts CR-014 section 12.6); 29: the reference run's size; 31: Word's
    # model of an inserted paragraph's mark; 32 and 33: a style's resets; 34: the hyperlink field
    assert differing == [28, 29, 31, 32, 33, 34]
    for n in (28, 29):
        assert [line.replace('"<commentReference>"{rStyle,sz}', '"<commentReference>"{rStyle}') for line in measured[n]] == mine[n]
    # what Word wrote, which this package now writes: a whole control a tracked deletion took, its content deleted
    # between Word's customXmlDelRange markers (15); a control around paragraphs held whole, tracked (43)
    assert measured[15] == mine[15] == [
        'P[Heading1|del] "HEAD15A " del:"HEAD15B HEAD15C"',
        "<customXmlDelRangeStart>",
        "SDT",
        "  <customXmlDelRangeEnd>",
        '  P[|del] del:"PLAIN15A PLAIN15B"',
        "  <customXmlDelRangeStart>",
        "<customXmlDelRangeEnd>",
        'P[center|] del:"BODY15A"{i} " BODY15B"{i}',
    ]
    assert measured[43][:3] == mine[43][:3] == ["<customXmlInsRangeStart>", "SDT", "  <customXmlInsRangeEnd>"]
    # and both the same: all but those six are the markup of Word's run, case 4's tracked Replace within one
    # paragraph among them since it puts its w:ins first (the user, in core-ts, 2026-10-23)
    assert len(measured) == 43
    assert measured[4] == mine[4]


def test_cr_014_section_12_6_a_deletion_takes_a_comments_reference_with_it_and_the_comment_goes_untracked_and_on_accept_all_reject_all_keeps_it():
    def made():
        package = create_package()
        paragraph = package.body.insert_paragraph("A B C D", location="End")
        paragraph.search("B C")[0].insert_comment("ON B C")
        return load(io.BytesIO(package.save()))

    def comments_of(data: bytes) -> list[str]:
        return [c.content for c in load(io.BytesIO(data)).body.get_comments()]

    # untracked: "C D" takes the reference after "C", and the comment goes (when the write ends here; core-ts on save)
    untracked = made()
    next(p for p in untracked.body.paragraphs if p.text.startswith("A")).search("C D")[0].delete()
    assert untracked.body.get_comments() == []
    assert comments_of(untracked.save()) == []
    # tracked: the reference run in the w:del with the text, as Word's (check 54 case 28)
    tracked = made()
    tracked.change_tracking_mode = "TrackAll"
    next(p for p in tracked.body.paragraphs if p.text.startswith("A")).search("C D")[0].delete()
    tracked.change_tracking_mode = "Off"
    assert re.search(
        r'<w:del [^>]*><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>'
        r'<w:r><w:delText xml:space="preserve"> D</w:delText>',
        tracked.main_document_part.xml.decode(),
    )
    saved = tracked.save()
    accepted = load(io.BytesIO(saved))
    revisions.accept_all(accepted)
    assert comments_of(accepted.save()) == []
    rejected = load(io.BytesIO(saved))
    revisions.reject_all(rejected)
    assert comments_of(rejected.save()) == ["ON B C"]
    # a replacement over the reference takes it too
    empty = made()
    paragraph = next(p for p in empty.body.paragraphs if p.text.startswith("A"))
    paragraph.search("A B C D")[0].insert_text("", location="Replace")
    assert paragraph.text == ""
    assert comments_of(empty.save()) == []


# ---------------------------------------------------------------------------
# Phase B's own: the refusals, the constructor, the members and the views
# ---------------------------------------------------------------------------


def test_words_own_refusals_inside_a_table_tracked_or_not_and_nothing_changes():
    """core-ts check 52 cases 9, 10, 22, 23, 34, 35, 44 and 45: Word refuses delete and Replace into or out of a table.

    Phase B refused them before its own refusal of every write across paragraphs; since Phase C they are what is
    left refused there, and a font, a split and a comment on such a range are Word's (cases 51, 59 and 81).
    """
    package = loaded("52-ranges.docx")
    before = package.body.text
    for mode in ("Off", "TrackAll"):
        package.change_tracking_mode = mode
        for start, end in (("HEAD9B", "T9R2C1A"), ("T10R1C1B", "BODY10A")):
            span = word_in(package, start).expand_to(word_in(package, end))
            for write in (
                span.delete,
                lambda span=span: span.insert_text("X", location="Replace"),
                lambda span=span: span.insert_ooxml(checks.PARAGRAPH_OOXML, location="Replace"),
                span.get_xml,
            ):
                with pytest.raises(SpanError) as refused:
                    write()
                assert refused.value.code == "range.partial_table", (mode, start)
                assert "inside a table" in str(refused.value)
    package.change_tracking_mode = "Off"
    assert package.body.text == before
    # a replacement of a range ending with a whole table: Word's refusal too (check 54 case 7)
    table_end = word_in(package, "HEAD21B").expand_to(word_in(package, "T21R1C1A").parent_table.get_range("Whole"))
    with pytest.raises(SpanError) as refused:
        table_end.insert_text("X", location="Replace")
    assert refused.value.code == "range.replaces_table"
    assert package.body.text == before


def test_a_range_across_paragraphs_is_checked_both_ends_in_one_story_the_end_not_before_the_start():
    package = loaded("52-ranges.docx")
    head = word_in(package, "HEAD3B")
    body3 = word_in(package, "BODY3A")
    made = Range(head.paragraph, head.start, body3.end, end_paragraph=body3.paragraph)
    assert made.text == "HEAD3B\rBODY3A"
    assert made == head.expand_to(body3)
    assert hash(made) == hash(head.expand_to(body3))
    assert len(made) == len("HEAD3B\rBODY3A")
    assert repr(made) == f"<Range 7:{body3.paragraph.ordinal}:6 'HEAD3B\\rBODY3A'>"
    assert made.to_dict() == {
        "address": head.paragraph.address,
        "text": "HEAD3B\rBODY3A",
        "start": 7,
        "end": 6,
        "end_address": body3.paragraph.address,
        "mark": False,
    }
    # a range within one paragraph reads as it did
    assert head.to_dict() == {"address": head.paragraph.address, "text": "HEAD3B", "start": 7, "end": 13}
    assert head.end_paragraph is head.paragraph and not head.spans_paragraphs
    with pytest.raises(SpanError) as order:
        Range(body3.paragraph, 0, 1, end_paragraph=head.paragraph)
    assert order.value.code == "range.order"
    note = next(p for p in package.main_document_part.footnotes_part.body.paragraphs if "NOTE12" in p.text)
    with pytest.raises(SpanError) as story:
        Range(head.paragraph, 0, 1, end_paragraph=note)
    assert story.value.code == "range.other_story"
    # expand_to takes either order (check 52's swapped reads) and leaves both ranges as they were
    assert body3.expand_to(head).text == "HEAD3B\rBODY3A"
    assert (head.text, body3.text) == ("HEAD3B", "BODY3A")


def test_a_comment_across_paragraphs_is_listed_by_a_range_on_any_of_them():
    """CR-013 section 2.5 item 1: a range on a comment's last paragraph, holding only its end marker, lists it."""
    package = loaded("52-ranges.docx")
    (comment,) = package.body.get_comments()
    # the comment runs from BODY1B into LIST1A
    assert [c.id for c in word_in(package, "LIST1A").get_comments()] == [comment.id]
    assert [c.id for c in word_in(package, "BODY1B").get_comments()] == [comment.id]
    assert word_in(package, "HEAD1A").get_comments() == [], "before it"
    # its end marker stands right after LIST1A, at 6: a range collapsed there lists it, both edges
    # inclusive as before; LIST1B, from 7, is after it
    list1 = word_in(package, "LIST1A").paragraph
    assert [c.id for c in Range(list1, 6, 6).get_comments()] == [comment.id]
    assert word_in(package, "LIST1B").get_comments() == []
    assert word_in(package, "PLAIN1B").get_comments() == []
    across = word_in(package, "HEAD1A").expand_to(word_in(package, "PLAIN1A"))
    assert [c.id for c in across.get_comments()] == [comment.id]


def test_a_comment_with_no_markers_in_the_body_has_no_range():
    package = create_package()
    paragraph = package.body.insert_paragraph("one two")
    comment = paragraph.search("two")[0].insert_comment("note")
    assert comment.get_range().text == "two"
    for item in list(paragraph.element.content):
        if type(item).__name__ in ("CommentRangeStart", "CommentRangeEnd"):
            paragraph.element.content.remove(item)
    # the reference run alone: collapsed where it stands, as get_ranges() reads it
    assert comment.get_range() is not None and comment.get_range().text == ""
    assert comment.get_ranges()[0].text == ""
    paragraph.element.content[:] = [
        item
        for item in paragraph.element.content
        if not any(
            (element_name(child) or "").endswith("}commentReference") for child in getattr(item, "content", None) or ()
        )
    ]
    assert comment.get_ranges() == []
    assert comment.get_range() is None


def test_check_48s_spans_from_expand_to_are_words_but_for_the_deleted_text_all_markup_shows():
    """CR-013 section 6.1: check 48's spans need expand_to, so they join in Phase B (core-ts 30b559e's snippet)."""
    measured = word("48-text-word365.json")
    body = load(OFFICE_JS / "48-text.docx").body
    differing = {}
    for name, theirs in measured["spans"].items():
        first, last = name.split("..")
        span = body.search(first, match_case=True)[0].expand_to(body.search(last, match_case=True)[0])
        if checks.vis(span.text) != theirs:
            differing[name] = checks.vis(span.text)
    # check 48 ran in All Markup: case 11's deleted text is in Word's span and not in Word's default view
    # (check 48b), which this package reads (Phase A's case 11)
    assert differing == {"REVA..REVB": "REVAINSERTEDTEXTREVB"}
    assert measured["spans"]["REVA..REVB"] == "REVADELETEDTEXTINSERTEDTEXTREVB"
    assert len(measured["spans"]) == 19


def test_the_style_rule_holds_within_one_paragraph_too():
    """core-ts check 53 cases 7 to 9: a range's character style where its runs share one, '' where they differ."""
    package = loaded("53-reads.docx")
    emphasis = word_in(package, "EMPH7B")
    assert (emphasis.style, emphasis.style_built_in) == ("Emphasis", "Emphasis")
    mixed = word_in(package, "PLAIN8A").expand_to(word_in(package, "EMPH8B"))
    assert (mixed.style, mixed.style_built_in) == ("", "Other")
    plain = word_in(package, "PLAIN9A")
    assert (plain.style, plain.style_built_in) == ("Normal", "Normal")


def test_a_row_controls_text_is_its_cells_a_tab_apart_and_its_whole_range_ends_with_two_tabs():
    """core-ts check 53 cases 15 to 17."""
    package = loaded("53-reads.docx")
    controls = {c.tag: c for c in package.body.content_controls}
    row = controls["c53-15"]
    assert row.text == "ROW15A\tROW15B"
    assert row.get_range("Whole").text == "ROW15A\tROW15B\t\t"
    assert row.get_range("Content").text == "ROW15A\tROW15B"
    ends_with_table = controls["c53-16"]
    assert ends_with_table.get_range("Whole").text.endswith("\t\t")


def test_range_content_controls_lists_those_it_overlaps_not_one_holding_it():
    """core-ts check 53 cases 19 to 21; get_content_controls filters them by type (G1)."""
    package = loaded("53-reads.docx")
    into = word_in(package, "PLAIN19B").expand_to(word_in(package, "CTRL19A"))
    assert [c.tag for c in into.content_controls] == ["c53-19"]
    inside = word_in(package, "CTRL20A").expand_to(word_in(package, "CTRL20B"))
    assert inside.content_controls == []
    over = word_in(package, "PLAIN21A").expand_to(word_in(package, "PLAIN21C"))
    assert [c.tag for c in over.content_controls] == ["c53-21"]
    assert [c.tag for c in over.get_content_controls(types=["RichText"])] == ["c53-21"]
    assert over.get_content_controls(types=["PlainText"]) == []


def test_the_returned_range_of_a_write_that_splits_paragraphs_spans_them():
    """Phase A's split writers (CR-013 section 11.1.6 item 1): the range across the paragraphs the text made."""
    package = create_package()
    paragraph = package.body.insert_paragraph("ab")
    out = paragraph.insert_text("x\ny", location="End")
    assert out.text == "x\ry"
    assert out.spans_paragraphs and [p.text for p in out.paragraphs] == ["abx", "y"]
    tracked = create_package()
    tracked.change_tracking_mode = "TrackAll"
    out = tracked.body.insert_paragraph("ab").insert_text("x\ry", location="End")
    assert out.text == "x\ry"


def test_body_and_table_get_range_refuse_a_location_they_do_not_take():
    package = loaded("52-ranges.docx")
    with pytest.raises(ContentError) as body_location:
        package.body.get_range("After")  # type: ignore[arg-type]
    assert body_location.value.code == "location.invalid"
    with pytest.raises(ContentError) as table_location:
        package.body.tables[0].get_range("Content")  # type: ignore[arg-type]
    assert table_location.value.code == "location.invalid"


def test_a_comment_whose_start_marker_stands_between_blocks_starts_at_the_next_paragraph():
    """Word's save of check 52 case 82: the w:commentRangeStart before a block control's first paragraph."""
    package = loaded("52-ranges-word365-accepted.docx")
    control = next(c for c in package.body.content_controls if c.tag == "c52-82")
    assert control.form == "block", "a block control whose content opens with the marker"
    (comment,) = [c for c in package.body.get_comments() if c.content == "NEWCOMMENT" and "82" in c.get_range().text]
    span = comment.get_range()
    assert span.text == "HEAD82A HEAD82B\rBODY82A BODY82B\rLIST82A LIST82B"
    assert span.paragraph.container is control.content, "the live list holding the paragraph"
    assert [c.id for c in word_in(package, "HEAD82A").get_comments()] == [comment.id]
