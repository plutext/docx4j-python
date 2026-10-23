"""Ranges across paragraphs, as Office JS's ``Word.Range`` is: CR-013 Phase B (the reads).

    .venv-fork/bin/python -m pytest tests/content/test_ranges.py

The port of docx4j-core-ts's ``test/ranges.test.mjs`` at ``53db5bc`` (its
CR-014), test for test, its strings kept: the five of phase 0 (check 52's file
and snippet agreeing, the ``TABLE_OOXML`` package, the two JSONs of this file,
Word's saves, ``accept_all`` on Word's save after the run against Word's Accept
All), the five of phase 1 (check 52's ``reads`` and ``report`` modes against
Word's, the tracked changes a range lists, a range from ``expand_to`` and the
views across paragraphs), the two of check 53, and the six of phase 2, the
writes, marked skip until CR-013 Phase C. core-ts runs check 52's and 53's
Script Lab snippets through its Office JS shim; this package has none, so
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

import hashlib
import importlib.util
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
        mine = dict(ours["reads"][kind])
        theirs = dict(theirs)
        our_marks, marks = mine.pop("marks"), theirs.pop("marks")
        # ^p within the range is core-ts's phase 3 (CR-013 Q10, not ported): no hit here, where Word
        # found each mark
        assert our_marks == [], kind
        assert len(marks) > 0 or kind.startswith("TC") or kind == "CELL", kind
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


def test_cr_014_phase_1_a_range_across_paragraphs_from_expand_to_reads_every_paragraph_the_writes_phase_2_does_not_bring_refuse():
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
    # every write across paragraphs is CR-013 Phase C's: refused before anything changes (core-ts
    # refused a control, a style, a link, a comment and insertOoxml in its phase 1; here every write)
    before = package.body.text

    def set_style() -> None:
        span.style = "Normal"

    def set_link() -> None:
        span.hyperlink = "https://example.com"

    def set_bold() -> None:
        span.font.bold = True

    for write in (
        span.insert_content_control,
        set_style,
        set_link,
        set_bold,
        lambda: span.insert_comment("X"),
        lambda: span.insert_ooxml("<w:p/>", location="After"),
        lambda: span.insert_text("X", location="End"),
        lambda: span.insert_paragraph("X", location="After"),
        span.delete,
        lambda: span.replace_text("3", "X"),
        lambda: setattr(span, "text", "X"),
        lambda: setattr(span, "style_built_in", "Normal"),
        lambda: setattr(span, "style_id", "Normal"),
        span.get_xml,  # the one w:p a fragment of it would be (core-ts's getOoxml)
        span.to_api_script,
    ):
        with pytest.raises(SpanError) as refused:
            write()
        assert refused.value.code == "range.spans_paragraphs"
        assert "Phase C" in str(refused.value)
    assert package.body.text == before
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
# phase 2, the writes (core-ts 53db5bc): CR-013 Phase C's
# ---------------------------------------------------------------------------

PHASE_C = "CR-013 Phase C: the writes of a range across paragraphs (core-ts 53db5bc, spanEdits.mts), not ported in Phase B"


@pytest.mark.skip(reason=PHASE_C)
def test_cr_014_phase_2_check_52s_writes_run_through_the_port_on_the_same_file_leave_each_region_as_word_365s_run_did():
    """core-ts: check 52's ``run`` mode against Word's run, cases 1 to 85, the differences listed (CR-013 section 6.1)."""


@pytest.mark.skip(reason=PHASE_C)
def test_cr_014_phase_2_accept_all_after_check_52s_writes_through_the_port_gives_word_365s_accept_all_after_its_run_region_for_region():
    """core-ts: Accept All after the run mode, saved and reloaded, against Word's report after Accept All."""


@pytest.mark.skip(reason=PHASE_C)
def test_cr_014_phase_2_the_markup_check_52s_writes_leave_through_the_port_is_word_365s_in_its_save_after_the_run_region_for_region():
    """core-ts: each region's markup, a line a paragraph, against Word's save after the run; cases 16, 29, 39, 56, 70 pinned."""


@pytest.mark.skip(reason=PHASE_C)
def test_cr_014_phase_2_delete_and_replace_across_paragraphs_and_the_range_a_write_leaves_reject_all_after_a_tracked_write_gives_back_what_was_there():
    """core-ts: delete and Replace join in the last paragraph's properties; a footnote leaves on save; Reject All."""


@pytest.mark.skip(reason=PHASE_C)
def test_cr_014_phase_2_whole_tables_a_cell_the_body_and_a_control_what_word_refuses_is_refused_before_anything_changes():
    """core-ts: whole tables, a cell, the body and a control deleted; Word's refusals inside a table (Phase B refuses them already, below)."""


@pytest.mark.skip(reason=PHASE_C)
def test_cr_014_phase_2_insert_paragraph_splits_a_paragraph_inside_a_link_into_two_links_and_refuses_a_split_inside_a_content_control_insert_text_at_the_ends_replace_text_a_paragraph_at_a_time():
    """core-ts: insertParagraph splitting a link, refused inside a control; insertText at the ends; replaceText a paragraph at a time."""


# ---------------------------------------------------------------------------
# Phase B's own: the refusals, the constructor, the members and the views
# ---------------------------------------------------------------------------


def test_words_own_refusals_inside_a_table_come_before_phase_cs_tracked_or_not_and_nothing_changes():
    """core-ts check 52 cases 9, 10, 22, 23, 34, 35, 44 and 45: Word refuses delete and Replace into or out of a table."""
    package = loaded("52-ranges.docx")
    before = package.body.text
    for mode in ("Off", "TrackAll"):
        package.change_tracking_mode = mode
        for start, end in (("HEAD9B", "T9R2C1A"), ("T10R1C1B", "BODY10A")):
            span = word_in(package, start).expand_to(word_in(package, end))
            for write in (span.delete, lambda span=span: span.insert_text("X", location="Replace")):
                with pytest.raises(SpanError) as refused:
                    write()
                assert refused.value.code == "range.partial_table", (mode, start)
                assert "inside a table" in str(refused.value)
    package.change_tracking_mode = "Off"
    # into a block content control from outside it: unmeasured, refused
    control = next(c for c in package.body.content_controls if c.tag == "c52-13")
    into = word_in(package, "BEGIN13").expand_to(control.paragraphs[0].get_range("Start"))
    with pytest.raises(SpanError) as refused:
        into.delete()
    assert refused.value.code == "range.partial_control"
    # a range holding a whole block control: refused under tracking, where Word's markup isn't measured
    around = word_in(package, "BEGIN26").expand_to(word_in(package, "END26"))
    package.change_tracking_mode = "TrackAll"
    with pytest.raises(SpanError) as refused:
        around.delete()
    assert refused.value.code == "range.not_trackable" and "sdt" in str(refused.value)
    package.change_tracking_mode = "Off"
    # untracked, the deletion itself is Phase C's
    with pytest.raises(SpanError) as refused:
        around.delete()
    assert refused.value.code == "range.spans_paragraphs"
    # a table's whole range, and a font or a paragraph on a range into a table, are Phase C's too
    for write in (
        package.body.tables[0].get_range().delete,
        lambda: setattr(word_in(package, "T51R1C1B").expand_to(word_in(package, "BODY51A")).font, "bold", True),
    ):
        with pytest.raises(SpanError) as refused:
            write()
        assert refused.value.code == "range.spans_paragraphs"
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
