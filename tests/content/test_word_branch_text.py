"""The Office JS readers read Word's branch of a kept ``mc:AlternateContent``: CR-011 Phase F (section 15).

    .venv-fork/bin/python -m pytest tests/content/test_word_branch_text.py -q

CR-011 Q4, revised 2026-10-22 (the user: "revise CR-011 Q4 so that
Paragraph.text, get_text and search read the branch Word shows"). Office JS's
``paragraph.text`` reads the branch Word shows of a run- or paragraph-level
``mc:AlternateContent``, never a Fallback's and never nothing, and
``body.search`` finds a word that only a Choice holds and none that only a
Fallback holds: docx4j-core-ts check 49 (Jason, Word 365, Script Lab,
2026-10-20; core-ts ``test/README.md`` check 49 and CR-012 section 12.3).

The probe is check 49's, ``tests/fixtures/mc-branches.docx`` (docx4j-ts-editor
``4e02089``). Word 365's answers are the report the snippet wrote into Jason's
save, ``mc-branches-check49-word365.docx`` (editor ``3189640``, ED-009 section
9.2 item 4), over the probe's first eight cases; cases 9 and 10 are Jason's
reading of the same day (ED-009 section 9.2 item 5). They are written out here
as :data:`WORD_365`, :data:`WORD_365_SEARCH` and :data:`WORD_365_AFTER_YY`.

What is held:

* each case's ``Paragraph.text`` is Word 365's, case 2's break U+000B among
  them since CR-013 Phase A (section 15 item 2, superseded 2026-10-23; until
  then pinned as ``\\n``);
* ``get_text`` in both views, ``Body.text``, ``search``, ``find`` and a range's
  text read the branch; a branch's insertions and deletions are filtered as
  the paragraph's own; a text box in a branch gives nothing (core-ts check
  42); the text model's ``office_offset`` agrees with the new text;
* every Word 365 save of docx4j's ``mc-fun`` probes reads as its probe does;
* an edit through a branch's text is refused (``address.read_only``) before
  anything changes, and an edit beside the element works, its offsets counting
  the branch's text;
* the internal readers of the offsets (a comment's range, a tracked change's
  range, a run control's span, ``Range.get_xml``) agree with the new text.
"""

from __future__ import annotations

import pytest
from conftest import FIXTURES

from docx4j_py import create_package, load
from docx4j_py.child import link_parents
from docx4j_py.model.content.errors import AddressError
from docx4j_py.model.content.range import Range
from docx4j_py.model.content.text_model import (
    office_as_model,
    segments_of,
    text_of_view,
)
from docx4j_py.model.text import extract, plain_text
from docx4j_py.traversal import text_of
from docx4j_py.wml import to_xml, wml

PROBE = FIXTURES / "mc-branches.docx"
MC_FUN = FIXTURES / "mc-fun"

#: Word 365's ``paragraph.text`` for each case's paragraph, by its index in the
#: probe's body (check 49's report, lines ``2:`` to ``16:``; cases 9 and 10,
#: ED-009 section 9.2 item 5). Case 2's break is U+000B, here as in Word since
#: CR-013 Phase A: :data:`WORD_365_BREAK`.
WORD_365 = {
    2: "The Earth is flat",
    4: "Before\tafter the tab\u000bafter the break",
    6: "unknown choice",
    8: "nothing to select",
    10: "A picture:  after it",
    12: "I prefer w14",
    14: "I prefer w16du",
    16: "bold ying plain ying",
    18: "I prefer w16sdtdh",
    20: "I prefer w16sdtfl",
}

#: Case 2 as Word 365 reads it: its ``w:br`` is U+000B (section 15 item 2).
WORD_365_BREAK = "Before\tafter the tab\u000bafter the break"

#: The report's ``search "...": n hit(s)`` lines.
WORD_365_SEARCH = {
    "yang": 0,
    "plain ying": 1,
    "bold ying": 1,
    "w16du": 2,
    "w14": 1,
    "after it": 1,
    "nothing to select": 1,
    "unknown choice": 1,
    "fallback shown": 0,
    "Fallback text": 0,
    "after the tab": 1,
    "round": 0,
    "flat": 1,
}

#: The report's ``after insertText('YY', 'After')`` line: the paragraph's text
#: after "YY" was inserted after the first "flat" hit.
WORD_365_AFTER_YY = "The Earth is flatYY"


def probe():
    return load(PROBE)


def body_xml(pkg) -> str:
    return to_xml(pkg.main_document_part.contents.body)


def built(body: str):
    """A package whose body is `body`, a ``w:body`` fragment."""
    pkg = create_package()
    pkg.main_document_part.contents.body = wml(body)
    link_parents(pkg.main_document_part.contents)
    return pkg


#: A paragraph with text before and after a run-level element, and one with a
#: paragraph-level element whose Choice holds an insertion and a deletion.
_AROUND = (
    "<w:body>"
    "<w:p><w:r><w:t xml:space=\"preserve\">head </w:t></w:r><w:r><mc:AlternateContent>"
    '<mc:Choice Requires="w14"><w:t>CHOICE</w:t></mc:Choice>'
    "<mc:Fallback><w:t>FALLBACK</w:t></mc:Fallback></mc:AlternateContent></w:r>"
    "<w:r><w:t xml:space=\"preserve\"> tail</w:t></w:r></w:p>"
    "<w:p><w:r><w:t>a</w:t></w:r><mc:AlternateContent><mc:Choice Requires=\"w14\">"
    "<w:r><w:t>b</w:t></w:r>"
    '<w:ins w:id="1" w:author="A" w:date="2026-10-22T00:00:00Z"><w:r><w:t>INS</w:t></w:r></w:ins>'
    '<w:del w:id="2" w:author="A" w:date="2026-10-22T00:00:00Z"><w:r><w:delText>DEL</w:delText></w:r></w:del>'
    "</mc:Choice><mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback></mc:AlternateContent>"
    "<w:r><w:t>c</w:t></w:r></w:p>"
    "<w:sectPr/></w:body>"
)


# ---------------------------------------------------------------------------
# the readers, against Word 365
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(("index", "word"), sorted(WORD_365.items()))
def test_each_cases_paragraph_text_is_word_365s(index: int, word: str) -> None:
    paragraph = probe().body.paragraphs[index]
    assert paragraph.text == word
    assert paragraph.get_text() == word
    assert paragraph.get_text(view="original") == word, "no tracked change: the two views agree"


def test_case_2s_break_is_u000b_as_office_js_gives_it() -> None:
    """Section 15 item 2, superseded 2026-10-23: CR-013 Phase A takes Office JS's characters.

    Until then this test was ``test_case_2s_break_is_a_newline_where_office_js_gives_u000b``,
    pinning the difference; now the break reads as Word 365's check 49 gave it,
    and docx4j's ``text_of`` keeps its ``\\n``.
    """
    paragraph = probe().body.paragraphs[4]
    assert paragraph.text == WORD_365_BREAK
    assert text_of(paragraph.element) == WORD_365_BREAK.replace("\u000b", "\n")


def test_the_case_labels_and_the_last_paragraph_are_as_before() -> None:
    texts = [p.text for p in probe().body.paragraphs]
    assert texts[1] == "Case 1: run-level, text branches"
    assert texts[15] == "Case 8: paragraph-level, runs in the branches"
    assert texts[-1] == ""
    assert len(texts) == 22


def test_paragraph_text_agrees_with_text_of_and_the_text_model() -> None:
    """Section 15: every reader of text takes the branch docx4j and core-ts take."""
    pkg = probe()
    paragraphs = pkg.body.paragraphs
    # the same branch; the characters differ only at case 2's break (CR-013 Phase A)
    assert [text_of_view(p.element, characters="docx4j") for p in paragraphs] == [
        text_of(p.element) for p in paragraphs
    ]
    assert [p.text.replace("\u000b", "\n") for p in paragraphs] == [text_of(p.element) for p in paragraphs]
    assert pkg.body.text == "\r".join(p.text for p in paragraphs)
    assert plain_text(extract(pkg)).rstrip("\n") == pkg.body.text.replace("\r", "\n").replace(
        "\u000b", "\n"
    ).rstrip("\n")


def test_search_finds_a_word_only_a_choice_holds_and_none_only_a_fallback_holds() -> None:
    body = probe().body
    assert [(h.paragraph.ordinal, h.start, h.end) for h in body.search("ying")] == [
        ("body/16", 5, 9),
        ("body/16", 16, 20),
    ]
    (flat,) = body.search("flat")
    assert (flat.paragraph.ordinal, flat.start, flat.end, flat.text) == ("body/2", 13, 17, "flat")
    assert flat.paragraph.text == "The Earth is flat", "check 49: the hit's paragraph"
    assert body.search("yang") == []
    assert body.search("round") == []


@pytest.mark.parametrize(("word", "hits"), sorted(WORD_365_SEARCH.items()))
def test_search_counts_are_word_365s(word: str, hits: int) -> None:
    """Check 49's search lines. Word's counts are its matchCase ones: under this package's default,
    case-insensitive search, ``unknown choice`` also hits the two labels' ``unknown Choice``."""
    assert len(probe().body.search(word, match_case=True)) == hits


def test_find_reports_a_branchs_hit_with_its_context() -> None:
    (hit,) = probe().body.find("flat")
    assert hit.to_dict() == {
        "address": "body/2",
        "ordinal": "body/2",
        "start": 13,
        "end": 17,
        "match": "flat",
        "snippet": "The Earth is flat",
        "before": "The Earth is ",
        "after": "",
    }
    body = probe().body
    assert body.range_of(hit).text == "flat"


def test_a_branchs_runs_are_filtered_as_the_paragraphs_own() -> None:
    """Section 15 item 1: the same deletions out in accepted, insertions out in original."""
    paragraph = built(_AROUND).body.paragraphs[1]
    assert paragraph.text == "abINSc"
    assert paragraph.get_text(view="accepted") == "abINSc"
    assert paragraph.get_text(view="original") == "abDELc"
    assert text_of(paragraph.element) == "abINSc"


def test_text_after_a_run_level_element_counts_its_text() -> None:
    pkg = built(_AROUND)
    paragraph = pkg.body.paragraphs[0]
    assert paragraph.text == "head CHOICE tail"
    (tail,) = paragraph.search("tail")
    assert (tail.start, tail.end) == (12, 16)
    assert pkg.body.search("FALLBACK") == []


def test_a_text_box_in_a_branch_gives_nothing() -> None:
    """core-ts check 42: a run-level text box, a wps Choice with its VML Fallback, adds nothing."""
    pkg = load(MC_FUN / "mc-txbx-run.docx")
    assert [p.text for p in pkg.body.paragraphs] == ["Text box, run level (Word's own form)", "RUN_PARA "]
    assert "RUN_CHOICE_BOX" in plain_text(extract(pkg)), "the text model reads the box, after its anchor"


@pytest.mark.parametrize(
    "stem",
    sorted(p.name.removesuffix("-word365.docx") for p in MC_FUN.glob("*-word365.docx")),
)
def test_each_mc_fun_probe_reads_as_word_365s_save_of_it(stem: str) -> None:
    """Word's save resolves each element to the branch it shows; the probe kept reads the same.

    Before Phase F, ``mc-run``, ``mc-mixed`` and ``mc-sdt-run-newest-first`` read the run-level
    branches as nothing (11 of 14 agreed).
    """
    kept = [p.text for p in load(MC_FUN / f"{stem}.docx").body.paragraphs]
    saved = [p.text for p in load(MC_FUN / f"{stem}-word365.docx").body.paragraphs]
    assert kept == saved


@pytest.mark.parametrize("view", ["accepted", "original"])
def test_each_runs_office_offset_is_where_its_text_is_in_get_text(view: str) -> None:
    """CR-011 12.6's invariant, a run read from a branch included (section 15 item 4)."""
    checked = 0
    for pkg in (probe(), built(_AROUND)):
        for block in extract(pkg, view=view).blocks:
            for run in block.runs:
                if not run.text:
                    continue
                assert run.office_offset is not None, (block.address, run.text)
                # agreement under CR-013 Q11's character map: Office JS's U+000B read as the model's \n
                office = office_as_model(segments_of(run.source.paragraph, view=view))
                assert office[run.office_offset : run.office_offset + len(run.text)] == run.text
                checked += 1
    assert checked == 25, "every run with text has its offset: the probe's 23 and the built two"


def test_a_branchs_run_has_its_offset() -> None:
    runs = extract(probe()).blocks[16].runs
    assert [(r.text, r.office_offset, r.bold) for r in runs] == [
        ("bold ying", 0, True),
        (" plain ying", 9, False),
    ]
    (run,) = extract(built(_AROUND)).blocks[0].runs
    assert (run.text, run.office_offset) == ("head CHOICE tail", 0)


def test_the_python_docx_facade_still_reads_python_docxs_text() -> None:
    """Q5: python-docx's ``Paragraph.text`` is its ``w:r`` and ``w:hyperlink`` children's text."""
    from docx4j_py.docx import Document

    document = Document(str(PROBE))
    texts = [p.text for p in document.paragraphs]
    assert texts[2] == "" and texts[12] == "I prefer " and texts[16] == ""


def test_the_facades_add_comment_on_a_run_holding_a_branch_is_refused() -> None:
    """``Document.add_comment`` anchors through the content API's offsets, so it follows the edit rule.

    The run whose only item is the element covers the branch's text (``w14``), so the comment is
    refused (``address.read_only``); before Phase F that run held no text and the comment went to
    offset 0, before "I prefer ". A comment on the run beside it is made.
    """
    from docx4j_py.docx import Document

    document = Document(str(PROBE))
    paragraph = document.paragraphs[12]
    with pytest.raises(AddressError) as caught:
        document.add_comment(paragraph.runs[1], "c")
    assert caught.value.code == "address.read_only"
    document.add_comment(paragraph.runs[0], "ok")
    (comment,) = paragraph.view.get_comments()
    assert [span.text for span in comment.get_ranges()] == ["I prefer "]  # Comment.get_range() is one Range since CR-013 Phase B (D2); get_ranges() the old list


# ---------------------------------------------------------------------------
# edits: refused through a branch's text, made beside it (section 15 item 3)
# ---------------------------------------------------------------------------


def _flat(body):
    return body.search("flat")[0]


_THROUGH = {
    "insert_text Replace": lambda b: _flat(b).insert_text("X", location="Replace"),
    "insert_text inside": lambda b: b.search("Earth")[0].insert_text("X", location="After"),
    "text": lambda b: setattr(_flat(b), "text", "X"),
    "delete": lambda b: _flat(b).delete(),
    "font": lambda b: setattr(_flat(b).font, "bold", True),
    "style": lambda b: setattr(_flat(b), "style", "Heading 1"),
    "insert_content_control": lambda b: _flat(b).insert_content_control(),
    "insert_comment": lambda b: _flat(b).insert_comment("note"),
    "hyperlink": lambda b: setattr(_flat(b), "hyperlink", "https://example.com/"),
    "insert_ooxml": lambda b: _flat(b).insert_ooxml("<w:p><w:r><w:t>Z</w:t></w:r></w:p>"),
    "replace_text on the range": lambda b: _flat(b).replace_text("flat", "round"),
    "replace_text on the paragraph": lambda b: b.paragraphs[16].replace_text("ying", "yang"),
    "replace_text on the body, a label's match first": lambda b: b.replace_text("w16du", "X"),
    "a caret inside": lambda b: Range(b.paragraphs[12], 10, 10).insert_text("X"),
    "a paragraph's font over a branch's text": lambda b: setattr(b.paragraphs[16].font, "italic", True),
    "a whole paragraph's range": lambda b: setattr(b.paragraphs[12].get_range().font, "bold", True),
    "insert_text Replace on the paragraph": lambda b: b.paragraphs[12].insert_text("X", location="Replace"),
    "insert_comment on the paragraph": lambda b: b.paragraphs[12].insert_comment("note"),
}


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("name", list(_THROUGH))
def test_an_edit_through_a_branchs_text_is_refused_before_anything_changes(name: str, tracked: bool) -> None:
    pkg = probe()
    if tracked:
        pkg.change_tracking_mode = "TrackAll"
    before = body_xml(pkg)
    parts = sorted(part.part_name.name for part in pkg.parts.values())
    changes = len(pkg.changes)
    with pytest.raises(AddressError) as caught:
        _THROUGH[name](pkg.body)
    assert caught.value.code == "address.read_only"
    assert "mc:AlternateContent" in str(caught.value)
    assert body_xml(pkg) == before
    assert sorted(part.part_name.name for part in pkg.parts.values()) == parts
    assert len(pkg.changes) == changes


def test_a_tracked_text_set_and_a_tracked_delete_are_refused_the_untracked_ones_take_the_element() -> None:
    """Section 15 item 3: no Word-consistent tracked form is measured, so tracked is refused."""
    pkg = probe()
    pkg.change_tracking_mode = "TrackAll"
    before = body_xml(pkg)
    for edit in (
        lambda: setattr(pkg.body.paragraphs[16], "text", "new"),
        lambda: pkg.body.paragraphs[2].delete(),
        lambda: pkg.body.clear(),
        lambda: pkg.body.insert_text("X", location="Replace"),
        lambda: pkg.body.paragraphs[12].insert_xml("<w:p><w:r><w:t>Z</w:t></w:r></w:p>", location="Replace"),
    ):
        with pytest.raises(AddressError) as caught:
            edit()
        assert caught.value.code == "address.read_only"
    assert body_xml(pkg) == before

    pkg = probe()
    case8 = pkg.body.paragraphs[16]
    case8.text = "new"
    assert case8.text == "new"
    assert "AlternateContent" not in to_xml(case8.element)
    pkg.body.paragraphs[2].delete()
    assert pkg.body.paragraphs[2].text == "Case 2: a tab and a break in the Choice"


def test_a_tracked_row_deletion_through_a_branchs_text_is_refused() -> None:
    pkg = built(
        "<w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        "<w:tr><w:tc><w:p><w:r><mc:AlternateContent><mc:Choice Requires=\"w14\"><w:t>CHOICE</w:t>"
        "</mc:Choice><mc:Fallback><w:t>FALLBACK</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p></w:tc></w:tr>"
        "<w:tr><w:tc><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl><w:sectPr/></w:body>"
    )
    pkg.change_tracking_mode = "TrackAll"
    before = body_xml(pkg)
    (table,) = pkg.body.tables
    with pytest.raises(AddressError):
        table.delete_rows(0, 1)
    with pytest.raises(AddressError):
        table.delete()
    assert body_xml(pkg) == before
    table.delete_rows(1, 1)
    assert "w:del" in to_xml(table.element)


def test_an_insertion_after_a_hit_at_the_elements_end_goes_beside_it_as_word_reads_it() -> None:
    """Check 49: ``insertText("YY", "After")`` on the first "flat" hit; the paragraph reads Word's text.

    Word puts the text in the branch; here it goes beside the element, every branch as it was, and
    the paragraph reads the same.
    """
    pkg = probe()
    flat = _flat(pkg.body)
    out = flat.insert_text("YY", location="After")
    paragraph = pkg.body.paragraphs[2]
    assert paragraph.text == WORD_365_AFTER_YY
    assert out.text == "YY" and flat.text == "flat"
    xml = to_xml(paragraph.element)
    assert "The Earth is round" in xml, "the Fallback as it was"
    assert xml.index("</mc:AlternateContent>") < xml.index("YY")


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_an_edit_beside_the_element_works_and_its_offsets_count_the_branch(tracked: bool) -> None:
    pkg = built(_AROUND)
    if tracked:
        pkg.change_tracking_mode = "TrackAll"
    paragraph = pkg.body.paragraphs[0]
    (tail,) = paragraph.search("tail")
    tail.font.bold = True
    tail.insert_text("!", location="After")
    Range(paragraph, 5, 5).insert_text("<")  # at the element's start
    Range(paragraph, 12, 12).insert_text(">")  # at its end, the "<" counted
    paragraph.insert_text("^", location="Start")
    assert paragraph.text == "^head <CHOICE> tail!"
    xml = to_xml(paragraph.element)
    assert "<w:t>CHOICE</w:t>" in xml and "<w:t>FALLBACK</w:t>" in xml, "both branches as they were"
    assert paragraph.search("CHOICE")[0].start == 7

    case8 = pkg.body.paragraphs[1]
    case8.insert_text("<", location="Start")
    case8.insert_text(">", location="End")
    assert case8.text == "<abINSc>"


def test_an_insertion_beside_a_paragraph_level_element_goes_outside_it() -> None:
    pkg = probe()
    case8 = pkg.body.paragraphs[16]
    case8.insert_text("^", location="Start")
    case8.insert_text("$", location="End")
    case8.insert_xml("<w:p><w:r><w:t>@</w:t></w:r></w:p>", location="End")
    assert case8.text == "^bold ying plain ying$@"
    branch = case8.element.content[1]
    assert type(branch).__name__ == "AlternateContent"
    assert [text_of(r) for r in branch.choice[0].any_element] == ["bold ying", " plain ying"]
    assert [p.text for p in pkg.body.paragraphs].count("^bold ying plain ying$@") == 1
    for run in case8.runs:
        assert run.parent is case8.element


def test_a_change_report_carries_the_branchs_text() -> None:
    pkg = probe()
    pkg.body.paragraphs[12].insert_text("!", location="End")
    report = pkg.changes[-1]
    assert (report.text_before, report.text_after) == ("I prefer w14", "I prefer w14!")


# ---------------------------------------------------------------------------
# the offsets' other readers
# ---------------------------------------------------------------------------


def test_a_comment_after_the_element_has_its_range() -> None:
    pkg = built(_AROUND)
    paragraph = pkg.body.paragraphs[0]
    comment = paragraph.search("tail")[0].insert_comment("note")
    (span,) = comment.get_ranges()  # Comment.get_range() is one Range since CR-013 Phase B (D2); get_ranges() the old list
    assert (span.start, span.end, span.text) == (12, 16, "tail")
    assert [c.content for c in paragraph.get_comments()] == ["note"]


def test_a_tracked_change_after_the_element_has_its_range() -> None:
    pkg = built(_AROUND)
    pkg.change_tracking_mode = "TrackAll"
    paragraph = pkg.body.paragraphs[0]
    paragraph.insert_text("!", location="End")
    (change,) = [c for c in paragraph.get_tracked_changes() if c.type == "Added"]
    span = change.get_range()
    assert (span.start, span.end, span.text) == (16, 17, "!")


def test_a_tracked_change_in_a_branch_has_its_range_there() -> None:
    paragraph = built(_AROUND).body.paragraphs[1]
    spans = {c.type: (c.get_range().start, c.get_range().end) for c in paragraph.get_tracked_changes()}
    assert spans == {"Added": (2, 5), "Deleted": (5, 5)}


def test_a_run_control_holding_a_branch_reads_and_extends_beside_it() -> None:
    pkg = built(
        "<w:body><w:p><w:r><w:t xml:space=\"preserve\">x </w:t></w:r><w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr>"
        "<w:sdtContent><w:r><mc:AlternateContent><mc:Choice Requires=\"w14\"><w:t>CHOICE</w:t></mc:Choice>"
        "<mc:Fallback><w:t>FALLBACK</w:t></mc:Fallback></mc:AlternateContent></w:r></w:sdtContent></w:sdt>"
        "<w:r><w:t xml:space=\"preserve\"> y</w:t></w:r></w:p><w:sectPr/></w:body>"
    )
    paragraph = pkg.body.paragraphs[0]
    assert paragraph.text == "x CHOICE y"
    (control,) = paragraph.content_controls
    assert control.text == "CHOICE"
    assert [(r.start, r.end) for r in control.search("CHOICE")] == [(2, 8)]
    assert control.get_range().text == "CHOICE"
    with pytest.raises(AddressError):
        control.insert_text("Z", location="Replace")
    control.insert_text("<", location="Start")
    control.insert_text(">", location="End")
    assert paragraph.text == "x <CHOICE> y"
    assert control.text == "<CHOICE>", "inside the control, beside the element"


def test_range_get_xml_keeps_an_element_its_span_holds_and_drops_one_outside_it() -> None:
    paragraph = probe().body.paragraphs[12]  # "I prefer " + the element's "w14"
    head = paragraph.search("I prefer ")[0].get_xml()
    assert "AlternateContent" not in head and "I prefer " in head
    both = Range(paragraph, 2, 12).get_xml()
    assert "AlternateContent" in both and "prefer " in both and "I " not in both
    inside = paragraph.search("w14")[0].get_xml()
    assert "AlternateContent" in inside and "prefer" not in inside
    assert paragraph.text == "I prefer w14", "the document is not touched"
