"""Office JS's characters in the readers, and the writers that take them back: CR-013 Phase A.

    .venv-fork/bin/python -m pytest tests/content/test_office_js_characters.py

The port of docx4j-core-ts's ``test/readers.test.mjs`` at ``9307bdb`` (its
CR-012 phase 4) for checks 48, 48b and 51, its strings kept, held to Word 365's
answers on the very files Word read (``tests/office_js/``, copied byte for byte
with a SHA-1 manifest, CR-013 Q8). What core-ts's file also holds and this
phase does not port --- text-box addresses, ``paragraphAt`` over notes and
comments, ``outline()`` on the text model --- is core-ts's CR-012 D3 to D5,
which this package has had since CR-011 or does not offer.

The rule (decided 2026-10-23, Q1: core-ts's whole set): a paragraph mark is
``\\r`` in ``Body.text`` and in every multi-paragraph text; a table is Word's
form; a ``w:br`` is U+000B, a page break U+000C, a column break U+000E, a
``w:cr`` U+000D; a note reference and a note's mark U+0002; a ``w:ptab`` ``0``;
the hyphens U+001E and U+001F; any ``w:sym`` ``(``; hidden text left out by the
text model's rule; an equation as Word draws it (``math_text.py``); a
paragraph's ``'Whole'`` range ends with its mark; and ``search`` takes Word's
Find codes, a picture as U+FFFC and an optional hyphen passed over.

Then Q2 and section 4.1's additions: every text writer maps each of those
characters back to the element it reads from, a comment's writer splits on
``\\r`` as well as ``\\n``, a bound control's write-back keeps docx4j's
characters, and the python-docx facade's writers follow python-docx 1.2.0.
Section 11's amendment (Word's own writers, core-ts checks 55 and 55b) moves two
things, each test marked: U+0002 and the other characters XML 1.0 excludes are
dropped, no longer refused (E3), and ``\\n`` and ``\\r`` through the splice
writers make paragraphs (E1), which ``test_word_writers.py`` holds.
"""

from __future__ import annotations

import hashlib
import json
import re

import pytest
from conftest import ROOT, reloaded

from docx4j_py import create_package, load
from docx4j_py.model.content.body import Body
from docx4j_py.model.content.errors import SpanError, TrackedChangeError
from docx4j_py.model.content.math_text import math_letter
from docx4j_py.model.content.text_model import (
    hidden_runs_of,
    office_as_model,
    segments_of,
    text_of_view,
)
from docx4j_py.model.text import extract
from docx4j_py.traversal import element_name, text_of
from docx4j_py.wml import to_xml

OFFICE_JS = ROOT / "tests" / "office_js"

#: The files Word 365 read, as the brief gives their SHA-1s (the user's share, ``__tmp/``): checks 48,
#: 48b and 51 (Phase A), and 52 and 53 with Word's two saves of check 52 (CR-013 Phase B, section 6.1).
SHARE_SHA1 = {
    "48-text.docx": "8dec79fdd0dffa36051d572c6093e6515321ac01",
    "48b-text.docx": "ca98e9c0bd6062233383ae9181cbd1893190f7dd",
    "51-text.docx": "7a6a0102051387e17d04e2eed87c46373c955005",
    "52-ranges.docx": "e014014dec6295b6abcd638f83f19e9b60e28997",
    "53-reads.docx": "aea48f83fd318b4260a73ac6ad9a99cd7a8f77f8",
    "52-ranges-word365.docx": "15f505306f8faf95f6946b47789849ccbf34603b",
    "52-ranges-word365-accepted.docx": "3d2a37367822fc323c0905e30b0931c6a3a4f58d",
    # CR-013 Phase C (Q8 as extended): check 54's file and Word's two saves (core-ts 87bbb38)
    "54-writes.docx": "98f548e35d47f28e87c38898a92f62e6b602b690",
    "54-writes-word365.docx": "b5e958f9204751488625ca5adafea3d605a89933",
    "54-writes-word365-accepted.docx": "5047a128aec36d2933d044ca6ae9b7ce62510daf",
    # CR-013 section 13.7: check 57's file and Word's three saves (core-ts 906d5e3)
    "57-fields-equations.docx": "5c8f4ea4d76685e126e2b58493bb33ead3d39513",
    "57-fields-equations-word365.docx": "e8ac302be7a8a7551da61c3a1306881ca0bb5fba",
    "57-fields-equations-word365-accepted.docx": "ee615f9e9a876a4aeddde7924da8c71ea6ec9d65",
    "57-fields-equations-word365-rejected.docx": "ad0937a076ff013b3e6e51c556899f465ebadef0",
}

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def unescape(text: str) -> str:
    """Word's JSON writes every character outside printable ASCII as ``\\uXXXX``, surrogates in pairs."""
    joined = re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), text)
    return joined.encode("utf-16", "surrogatepass").decode("utf-16")


def word(name: str) -> dict:
    return json.loads((OFFICE_JS / name).read_text(encoding="utf-8"))


def views() -> list[dict]:
    """Check 48b's five runs on check 48's file (A Simple Markup, B hidden shown, C All Markup, D No Markup, E Original)."""
    text = (OFFICE_JS / "48-text-views-word365.txt").read_text(encoding="utf-8")
    out = []
    at = text.find('{\n "label"')
    while at >= 0:
        depth = 0
        end = at
        for end in range(at, len(text)):
            if text[end] == "{":
                depth += 1
            elif text[end] == "}":
                depth -= 1
                if depth == 0:
                    break
        out.append(json.loads(text[at : end + 1]))
        at = text.find('{\n "label"', at + 1)
    return out


def fixture(name: str):
    return load(OFFICE_JS / name)


def mini(body_xml: str):
    """A package of one body, through a save and a load, as core-ts's ``mini``."""
    made = create_package()
    made.body.content.clear()
    made.body.insert_xml(body_xml)
    return reloaded(made)


def xml_of(package) -> str:
    return to_xml(package.main_document_part.contents)


# ---------------------------------------------------------------------------
# the fixtures (Q8)
# ---------------------------------------------------------------------------


def test_every_fixture_is_the_manifests_and_the_files_word_read():
    lines = (OFFICE_JS / "SHA1SUMS").read_text(encoding="utf-8").splitlines()
    manifest = dict(reversed(line.split("  ", 1)) for line in lines if line.strip())
    # Phase A's seven, Phase B's eight (checks 52 and 53, Word's saves of 52 and Word's four JSONs)
    # and Phase C's ten (check 54's file, Word's two saves and its two JSONs; check 57's file, Word's
    # three saves and its run JSON, CR-013 section 13.7); and Word's Reject All of check 54 (docx4j
    # e55a83b8e, CR-010 section 13.3a)
    assert len(manifest) == 26
    for name, digest in manifest.items():
        assert hashlib.sha1((OFFICE_JS / name).read_bytes()).hexdigest() == digest, name
    for name, digest in SHARE_SHA1.items():
        assert manifest[name] == digest, f"{name} is the file Word 365 read"


# ---------------------------------------------------------------------------
# Paragraph.text and Body.text (core-ts D6, CR-012 section 6.2)
# ---------------------------------------------------------------------------


def text_after_label(paragraphs, n: int) -> str:
    """A case's paragraph's text after its ``Case n: ... .`` label."""
    found = next(p for p in paragraphs if p.text.startswith(f"Case {n}:"))
    return re.sub(r"^Case \d+: [^.]*\. ", "", found.text)


def test_check_48_and_48b_paragraph_text_is_office_jss_in_words_default_view_paragraph_for_paragraph():
    simple = views()[0]
    paragraphs = fixture("48-text.docx").body.paragraphs
    assert len(paragraphs) == len(simple["paragraphs"])
    for index, (paragraph, measured) in enumerate(zip(paragraphs, simple["paragraphs"], strict=True)):
        assert paragraph.text == unescape(measured["text"]), f"check 48, paragraph {index}"
    # the characters, as check 48 measured them

    def text(n: int) -> str:
        return text_after_label(paragraphs, n)

    assert re.sub(r"^.*?(FOOTA)", r"\1", text(1)) == "FOOTA\u0002FOOTB and ENDA\u0002ENDB"
    assert text(3) == "LINEA\u000bLINEB"
    assert text(4) == "PAGEA\u000cPAGEB"
    assert text(5) == "COLA\u000eCOLB"
    assert text(6) == "CRA\rCRB"
    assert text(7) == "TABA\tTABB0RIGHTB"
    assert text(8) == "HYPHA\u001eHYPHB\u001fHYPHC"
    assert text(9) == "SYMA(SYMB"
    assert text(10)[-8:] == "HIDAHIDB", "hidden text left out"
    assert text(11)[-20:] == "REVAINSERTEDTEXTREVB", "the accepted text: Simple Markup"
    assert text(16)[-8:] == "PICAPICB"
    assert text(17).endswith("MATHA \U0001d465=1 MATHB"), "the equation as Word draws it, its x Unicode's mathematical italic"
    # 48b's take 2, in Word's defaults: every symbol "(", a run hidden through its character style left out
    take2 = word("48b-text-word365.json")
    b = fixture("48b-text.docx").body.paragraphs
    assert len(b) == len(take2["paragraphs"])
    for index, (paragraph, measured) in enumerate(zip(b, take2["paragraphs"], strict=True)):
        assert paragraph.text == unescape(measured["text"]), f"check 48b, paragraph {index}"


def test_check_48_a_whole_range_ends_with_its_mark_and_body_text_is_words():
    measured = word("48-text-word365.json")
    package = fixture("48-text.docx")

    def simple(text: str) -> str:
        # check 48 ran in All Markup: only case 11 reads its deleted text there, and Simple Markup is the rule
        return unescape(text).replace("REVADELETEDTEXTINSERTEDTEXTREVB", "REVAINSERTEDTEXTREVB")

    for index, paragraph in enumerate(package.body.paragraphs):
        assert paragraph.get_range("Whole").text == simple(measured["paragraphs"][index]["whole"]), index
        assert paragraph.get_range("Content").text == paragraph.text
    # its table's row ended \t\t in All Markup with hidden text shown; Word's default view ends a
    # row \r\n (core-ts checks 30, 31 and 51)
    assert package.body.text == simple(measured["bodyText"]).replace(
        "CELLA\tCELLB\t\tAFTERTABLE", "CELLA\tCELLB\r\nAFTERTABLE"
    )
    assert "REVADELETEDTEXTREVB" in package.body.get_text(view="original")
    # two paragraphs in a cell: the first's mark \r, the last's the cell's end
    two = mini(
        f'<w:tbl xmlns:w="{W}"><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>'
        "<w:tc><w:p/></w:tc></w:tr></w:tbl>"
    )
    two.body.insert_paragraph("z")
    assert two.body.text == "a\rb\t\r\nz"
    assert [p.get_range().text for p in two.body.paragraphs] == ["a\r", "b\t", "\t", "z\r"]


def test_check_51_cells_controls_comments_and_notes_of_several_paragraphs_read_as_word_365s():
    measured = word("51-text-word365.json")
    package = fixture("51-text.docx")
    body = package.body
    # every paragraph, its whole range, and body.text: the equations' letters, minus signs and structures among them
    assert len(body.paragraphs) == len(measured["paragraphs"])
    for index, paragraph in enumerate(body.paragraphs):
        assert paragraph.text == unescape(measured["paragraphs"][index]["text"]), f"paragraph {index}"
        assert paragraph.get_range("Whole").text == unescape(measured["paragraphs"][index]["whole"]), index
    assert body.text == unescape(measured["bodyText"])
    # the joiner is the paragraph mark, \r, as body.text has it
    (table,) = body.tables
    assert table.values == [[unescape(v) for v in row] for row in measured["tables"][0]["values"]]
    for index, cell in enumerate(table.rows[0].cells):
        cell_measured = measured["tables"][0]["cells"][index]
        assert [cell.value, cell.body.text, [p.text for p in cell.paragraphs]] == [
            unescape(cell_measured["value"]),
            unescape(cell_measured["bodyText"]),
            [unescape(p) for p in cell_measured["paragraphs"]],
        ]
        last = cell.paragraphs[-1]
        assert unescape(cell_measured["bodyWhole"]).endswith(last.get_range().text), "a cell's last mark is \\t"
    assert table.text == "CELLP1\nCELLP2\nCELLP3\nOTHER", "Table.text stays docx4j's, a line per paragraph"
    block, inline = body.content_controls
    assert block.text == unescape(measured["contentControls"][0]["text"])
    assert block.docx4j_text == "CCP1\nCCP2\nCCP3", "docx4j's, for a bound control's write-back"
    assert inline.text == unescape(measured["contentControls"][1]["text"])
    assert [inline.get_range("Whole").text, inline.get_range("Content").text] == [
        measured["contentControls"][1]["whole"],
        measured["contentControls"][1]["content"],
    ]
    # a block control of several paragraphs: one range over the control since CR-013 Phase B (it was
    # its first paragraph's, "CCP1\r", until then), Word's "CCP1\rCCP2\rCCP3\r"
    assert unescape(measured["contentControls"][0]["whole"]) == "CCP1\rCCP2\rCCP3\r"
    assert block.get_range("Whole").text == unescape(measured["contentControls"][0]["whole"])
    assert block.get_range("Content").text == "CCP1\rCCP2\rCCP3"
    assert body.get_comments()[0].content == unescape(measured["comments"][0]["content"])
    # a note's own mark reads U+0002, as its reference does
    for part, notes in (
        (package.main_document_part.footnotes_part, measured["footnotes"]),
        (package.main_document_part.endnotes_part, measured["endnotes"]),
    ):
        note = next(n for n in (part.contents.footnote if hasattr(part.contents, "footnote") else part.contents.endnote) if n.id == 1)
        note_body = Body(part, note, prefix="note:1", package=package)
        assert [p.text for p in note_body.paragraphs] == [unescape(p) for p in notes[0]["paragraphs"]]
        assert note_body.text == unescape(notes[0]["bodyText"])
        whole = unescape(notes[0]["bodyWhole"])
        assert note_body.paragraphs[-1].get_range("Whole").text == "\r".join(whole.split("\r")[-2:])


# ---------------------------------------------------------------------------
# hidden text (core-ts check 48b; the text model's rule, REQ-507)
# ---------------------------------------------------------------------------

HIDDEN = (
    f'<w:p xmlns:w="{W}"><w:r><w:t xml:space="preserve">see </w:t></w:r>'
    "<w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r>"
    '<w:r><w:t xml:space="preserve"> here</w:t></w:r></w:p>'
)


def test_hidden_text_is_left_out_an_edit_across_it_takes_it_and_an_insertion_beside_it_does_not_go_into_it():
    package = mini(HIDDEN)
    p = package.body.paragraphs[0]
    assert p.text == "see  here"
    assert p.get_text(view="original") == "see  here"
    assert p.search("HIDDEN") == [], "Word's default view: a hidden word isn't found"
    assert len(p.search("see  here")) == 1, "and a query matches across it"
    p.insert_text("X", location="End")
    p.get_range("Start").insert_text("Y", location="Start")
    p.search(" here")[0].insert_text("!", location="Start")
    xml = xml_of(package)
    assert re.search(r"<w:t(?: xml:space=\"preserve\")?>Ysee !?</w:t>", xml), xml
    assert re.search(r"<w:vanish/></w:rPr><w:t>HIDDEN</w:t>", xml), "the hidden run as it was"
    p.search("see ! here")[0].delete()
    assert p.text == "YX"
    assert "HIDDEN" not in xml_of(package), "the hidden text inside the span went with it"


def test_a_run_hidden_through_its_character_style_is_left_out_and_a_style_added_later_is_seen():
    package = fixture("48b-text.docx")
    paragraph = next(p for p in package.body.paragraphs if p.text.startswith("Case 20:"))
    assert paragraph.text.endswith("HIDSAHIDSB")
    assert "STYLEHIDDEN" in text_of(paragraph.element), "docx4j's text_of reads it"
    resolver = package.get_property_resolver()
    hidden = [run for run in paragraph.runs if hidden_runs_of(resolver, paragraph.element)(run)]
    assert [text_of(run) for run in hidden] == ["STYLEHIDDEN"]


def test_a_tracked_deletion_across_hidden_text_puts_it_in_the_w_del():
    package = mini(HIDDEN)
    package.change_tracking_mode = "TrackAll"
    p = package.body.paragraphs[0]
    p.search("see  here")[0].delete()
    xml = xml_of(package)
    assert "<w:delText>HIDDEN</w:delText>" in xml
    assert p.text == ""
    package.body.reject_all()
    assert text_of(package.body.paragraphs[0].element) == "see HIDDEN here"


def test_setting_a_paragraphs_text_drops_its_hidden_runs_untracked_and_tracked():
    """CR-013 risk 7: ``paragraph.text = paragraph.text`` drops hidden runs, as core-ts's does."""
    package = mini(HIDDEN)
    p = package.body.paragraphs[0]
    p.text = p.text
    assert "HIDDEN" not in xml_of(package)
    tracked = mini(HIDDEN)
    tracked.change_tracking_mode = "TrackAll"
    q = tracked.body.paragraphs[0]
    q.text = "new"
    assert "<w:delText>HIDDEN</w:delText>" in xml_of(tracked), "deleted with the rest, whole"


# ---------------------------------------------------------------------------
# search (core-ts CR-012 section 6.2 question 3)
# ---------------------------------------------------------------------------


def test_search_words_own_queries_give_words_hits():
    simple = views()[0]
    take2 = word("48b-text-word365.json")
    check48 = word("48-text-word365.json")
    for label, name, searches, view_dependent in (
        ("48b run A", "48-text.docx", simple["searches"], []),
        ("48b take 2", "48b-text.docx", take2["searches"], []),
        # check 48 ran in All Markup with hidden text shown: the hidden and deleted words read otherwise there
        (
            "check 48",
            "48-text.docx",
            check48["searches"],
            [
                "HIDAHIDB",
                "HIDDENTEXT",
                "HIDAHIDDENTEXTHIDB",
                "DELETEDTEXT",
                "REVAINSERTEDTEXTREVB",
                "REVADELETEDTEXTINSERTEDTEXTREVB",
            ],
        ),
    ):
        package = fixture(name)
        differ = []
        for query, hits in searches.items():
            mine = [r.text for r in package.body.search(query)]
            if mine != [unescape(h) for h in hits]:
                differ.append(query)
        assert differ == view_dependent, label


def test_search_find_codes_tell_the_two_references_apart_and_a_range_reads_its_own_span():
    package = fixture("48-text.docx")
    assert package.body.search("FOOTA^eFOOTB") == []
    assert package.body.search("ENDA^fENDB") == []
    p = next(q for q in package.body.paragraphs if "HYPHA" in q.text)
    span = p.search("HYPHB^-HYPHC")[0]
    assert span.text == "HYPHB\u001fHYPHC"
    assert len(span.search("HYPHBHYPHC")) == 1
    assert span.search("HYPHA") == []
    assert [h.match for h in package.body.find("LINEA^lLINEB")] == ["LINEA\u000bLINEB"]
    assert [r.text for r in package.body.search("^^")] == [], "a caret, of which the file has none"


# ---------------------------------------------------------------------------
# equations (math_text.py, core-ts mathText.mts)
# ---------------------------------------------------------------------------

MATH = (
    f'<m:oMath xmlns:m="{M}"><m:r><m:t>x</m:t></m:r><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>+1</m:t></m:r></m:oMath>'
)


def equation_package():
    return mini(
        f'<w:p xmlns:w="{W}"><w:r><w:t xml:space="preserve">a </w:t></w:r>{MATH}'
        '<w:r><w:t xml:space="preserve"> b</w:t></w:r></w:p>'
    )


def test_math_letter_draws_as_word_does():
    assert math_letter("x", None, None) == "\U0001d465"
    assert math_letter("h", None, None) == "ℎ", "italic h is the BMP's Planck constant"
    assert math_letter("2", None, None) == "2"
    assert math_letter("x", "p", None) == "x"
    assert math_letter("x", "b", None) == "\U0001d431"
    assert math_letter("R", None, "double-struck") == "ℝ"
    assert math_letter("α", None, None) == "\U0001d6fc", "small Greek italic"
    assert math_letter("Δ", None, None) == "\U0001d6e5", "capital Greek italic too (check 51)"
    assert math_letter("2", "bi", None) == "\U0001d7d0", "bold digits under bold italic (check 51)"


def test_an_equation_reads_as_word_draws_it_and_an_edit_takes_it_whole_and_untracked_or_not_at_all():
    package = equation_package()
    p = package.body.paragraphs[0]
    assert p.text == "a \U0001d465+1 b"
    with pytest.raises(SpanError) as raised:
        p.search("+1 b")[0].delete()
    assert raised.value.code == "range.cuts_equation"
    with pytest.raises(SpanError):
        p.splice(3, 3, "z")
    p.search("b")[0].font.bold = True
    p.get_range("Whole").font.italic = True
    assert re.search(r"<m:oMath><m:r><m:t>x</m:t></m:r>", xml_of(package)), "a format leaves the math runs' own"
    p.search("a \U0001d465+1 ")[0].delete()
    assert p.text == "b"
    assert "oMath" not in xml_of(package), "taken whole"
    # tracked, a deletion takes the equation whole in Word's form, a w:del in each math run (CR-013 section 13.7,
    # D2; docx4j-core-ts check 57), where Phase A refused it (tracking.equation)
    tracked = equation_package()
    tracked.change_tracking_mode = "TrackAll"
    tracked.body.paragraphs[0].get_range("Content").delete()
    assert re.search(r"<m:oMath><m:r><w:del ", xml_of(tracked))
    assert tracked.body.paragraphs[0].text == ""
    cleared = equation_package()
    cleared.change_tracking_mode = "TrackAll"
    cleared.body.clear()
    assert re.search(r"<m:oMath><m:r><w:del ", xml_of(cleared))
    # a tracked replacement taking one is still refused, before anything changes (unmeasured where Word puts the text)
    replaced = equation_package()
    replaced.change_tracking_mode = "TrackAll"
    before = xml_of(replaced)
    with pytest.raises(TrackedChangeError) as raised:
        replaced.body.paragraphs[0].insert_text("z", location="Replace")
    assert raised.value.code == "tracking.equation"
    assert xml_of(replaced) == before, "refused before anything changed"
    # a range's XML keeps the equation whole where the span holds part of it
    assert "oMath" in equation_package().body.paragraphs[0].search("+1 b")[0].get_xml()


# ---------------------------------------------------------------------------
# the office_offset invariant under the character map (Q11)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", ["48-text.docx", "48b-text.docx", "51-text.docx"])
@pytest.mark.parametrize("view", ["accepted", "original"])
def test_each_runs_office_offset_agrees_under_the_character_map(name, view):
    package = fixture(name)
    resolver = package.get_property_resolver()
    checked = 0
    for block in extract(package, view=view).blocks:
        for run in block.runs:
            if run.office_offset is None:
                continue
            p = run.source.paragraph
            office = office_as_model(segments_of(p, view=view, hidden=hidden_runs_of(resolver, p)))
            assert office[run.office_offset : run.office_offset + len(run.text)] == run.text, block.address
            checked += 1
    assert checked > 20


# ---------------------------------------------------------------------------
# the writers (Q2): each character the element it reads from
# ---------------------------------------------------------------------------

#: Each of Office JS's characters a writer maps, the element it becomes, and what it reads back as.
MAPPED = [
    ("\u000b", "<w:br/>", "\u000b"),
    ("\u000c", '<w:br w:type="page"/>', "\u000c"),
    ("\u000e", '<w:br w:type="column"/>', "\u000e"),
    ("\r", "<w:cr/>", "\r"),
    ("\u001e", "<w:noBreakHyphen/>", "\u001e"),
    ("\u001f", "<w:softHyphen/>", "\u001f"),
    ("\t", "<w:tab/>", "\t"),
    ("\n", "<w:br/>", "\u000b"),
]


def _paragraph_text(package):
    return package.body.paragraphs[0]


WRITERS = {
    "Body.insert_paragraph": lambda pkg, text: pkg.body.insert_paragraph(text, location="Start"),
    "Paragraph.text": lambda pkg, text: setattr(_paragraph_text(pkg), "text", text),
    "Range.insert_text": lambda pkg, text: _paragraph_text(pkg).get_range("End").insert_text(text, location="End"),
    "Paragraph.insert_text Replace": lambda pkg, text: _paragraph_text(pkg).insert_text(text, location="Replace"),
    "Body.replace_text": lambda pkg, text: pkg.body.replace_text("seed", text),
    "Paragraph.insert_paragraph": lambda pkg, text: _paragraph_text(pkg).insert_paragraph(text, location="Before"),
}


def _seeded(tracked: bool):
    package = create_package()
    package.body.insert_paragraph("seed")
    if tracked:
        package.change_tracking_mode = "TrackAll"
    return package


#: The writers through ``Paragraph.splice``: ``\\n`` and ``\\r`` make a paragraph there since CR-013
#: section 11 (E1), which ``test_word_writers.py`` holds; the rest keep Q2's map.
SPLICE_WRITERS = frozenset({"Range.insert_text", "Paragraph.insert_text Replace", "Body.replace_text"})

WRITER_CASES = [
    pytest.param(writer, char, element, reads, id=f"{char!r}-{writer}")
    for writer in sorted(WRITERS)
    for char, element, reads in MAPPED
    if not (writer in SPLICE_WRITERS and char in "\r\n")
]


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize(("writer", "char", "element", "reads"), WRITER_CASES)
def test_each_writer_writes_each_character_as_the_element_it_reads_from(writer, char, element, reads, tracked):
    package = _seeded(tracked)
    WRITERS[writer](package, f"a{char}b")
    xml = xml_of(package)
    assert element in xml, (writer, xml)
    assert char not in "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml)) or char == "\n", "no raw character in a w:t"
    texts = [p.text for p in package.body.paragraphs]
    assert f"a{reads}b" in "".join(texts), texts
    again = reloaded(package)
    assert [p.text for p in again.body.paragraphs] == texts, "it saves, and reads back as written"


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_the_cell_comment_control_and_table_writers_map_too(tracked):
    package = _seeded(tracked)
    package.body.insert_table(1, 2, values=[["c\u000bd", "e\u001ef"]])
    table = package.body.tables[0]
    table.add_rows(1, values=[["g\u000ch", "i\u001fj"]])
    assert table.values == [["c\u000bd", "e\u001ef"], ["g\u000ch", "i\u001fj"]]
    table.cell(0, 0).value = "k\u000el"
    assert table.cell(0, 0).value == "k\u000el"
    comment = package.body.paragraphs[0].insert_comment("one\u000btwo")
    assert comment.content == "one\u000btwo"
    comment.content = "three\u000cfour\rfive"
    assert comment.content == "three\u000cfour\rfive", "\\r is a paragraph mark in a comment"
    assert len(comment.paragraphs) == 2
    package.body.insert_xml(
        f'<w:p xmlns:w="{W}"><w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>'
        "<w:r><w:t>in</w:t></w:r></w:sdtContent></w:sdt></w:p>"
    )
    control = package.body.content_controls[-1]
    control.insert_text("x\u000by", location="End")
    assert control.text == "inx\u000by"
    again = reloaded(package)
    assert again.body.tables[0].values == table.values
    assert again.body.get_comments()[0].content == "three\u000cfour\rfive"


def test_a_comments_writer_splits_on_r_as_well_as_n():
    """Section 4.1: ``"one\\rtwo\\nthree"`` saved two paragraphs, ``one&#13;two`` and ``three``, before Phase A."""
    package = _seeded(False)
    comment = package.body.paragraphs[0].insert_comment("x")
    comment.content = "one\rtwo\nthree\r\nfour"
    assert comment.content == "one\rtwo\rthree\rfour"
    assert len(comment.paragraphs) == 4
    assert "&#13;" not in to_xml(comment.element)
    assert reloaded(package).body.get_comments()[0].content == "one\rtwo\rthree\rfour"


@pytest.mark.parametrize("char", ["\u0002", "\u0001", "\u0000", "\u0010", "\u001d", "￾", "\ud800"])
def test_the_other_non_xml_characters_are_dropped_as_word_drops_them(char):
    """CR-013 section 11, E3 ("Drop, as Word does"): Phase A refused them (``text.not_writable``); now dropped.

    Word 365 dropped U+0002, U+0001 and U+0007 through every writer and refused
    none (core-ts check 55 at ``a923aeb``, cases 10 to 12, 24 to 26, 38 to 40).
    """
    for write, read in (
        *((lambda pkg, w=w: w(pkg, f"a{char}b"), lambda pkg: "".join(p.text for p in pkg.body.paragraphs))
          for w in WRITERS.values()),
        (lambda pkg: pkg.body.paragraphs[0].insert_comment(f"a{char}b"), lambda pkg: pkg.body.get_comments()[0].content),
        (lambda pkg: pkg.body.insert_table(1, 1, values=[[f"a{char}b"]]), lambda pkg: pkg.body.tables[0].values[0][0]),
        (lambda pkg: pkg.body.insert_text(f"a{char}b", location="Replace"), lambda pkg: pkg.body.text),
    ):
        package = _seeded(False)
        write(package)
        assert char not in xml_of(package)
        assert "ab" in read(package), "dropped, the text either side kept"
        assert char not in read(reloaded(package)), "it saves, and reads back without it"


def test_a_bound_controls_write_back_keeps_docx4js_characters():
    """Section 4.1: a break ``\\n`` and a block control's paragraphs a ``\\n`` apart, as before Phase A (measured at e7e54e6)."""
    greeting = '<greeting xmlns="http://example.com/g"><to>World</to></greeting>'
    namespaces = "xmlns:ns0='http://example.com/g'"
    path = "/ns0:greeting[1]/ns0:to[1]"
    for xml, expected in (
        (
            (
                f'<w:p xmlns:w="{W}"><w:r><w:t xml:space="preserve">at </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>'
                "<w:text/></w:sdtPr><w:sdtContent><w:r><w:t>line one</w:t><w:br/><w:t>line two</w:t></w:r>"
                "</w:sdtContent></w:sdt></w:p>"
            ),
            "line one\nline two",
        ),
        (
            (
                f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="8"/></w:sdtPr><w:sdtContent>'
                "<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t><w:br/><w:t>third</w:t></w:r></w:p>"
                "</w:sdtContent></w:sdt>"
            ),
            "first\nsecond\nthird",
        ),
    ):
        package = create_package()
        package.body.insert_xml(xml)
        control = package.body.content_controls[0]
        part = package.custom_xml_parts.add(greeting)
        control.xml_mapping.set_mapping(path, namespaces, part)
        assert "\u000b" in control.text, "Office JS's text holds the U+000B"
        package.custom_xml_parts.update_from_content_controls()
        assert part.select_single_node(path).text == expected
        assert reloaded(package).custom_xml_parts.items[0].select_single_node(path).text == expected


# ---------------------------------------------------------------------------
# the python-docx facade: python-docx 1.2.0's behaviour, measured 2026-10-23
# ---------------------------------------------------------------------------


def test_the_facades_writers_are_python_docx_1_2_0s():
    from docx4j_py.docx import Document

    document = Document()
    cases = {
        "add_paragraph": lambda text: document.add_paragraph(text),
        "add_run": lambda text: _run_of(document, text),
        "run.text": lambda text: _set_run_text(document, text),
        "paragraph.text": lambda text: _set_paragraph_text(document, text),
    }
    for name, write in cases.items():
        for text, breaks in (("a\rb", 1), ("a\nb", 1), ("a\r\nb", 2)):
            paragraph = write(text)
            xml = to_xml(paragraph._p)
            assert xml.count("<w:br/>") == breaks, (name, text, xml)
            assert "&#13;" not in xml and "<w:cr/>" not in xml, (name, text)
            assert paragraph.text == "a" + "\n" * breaks + "b"
        for char in ("\u000b", "\u0002", "\u000c", "\u001e"):
            with pytest.raises(ValueError):
                write(f"a{char}b")
    cell = document.add_table(1, 1).cell(0, 0)
    cell.text = "a\rb\nc"
    assert cell.text == "a\nb\nc"
    target = document.add_paragraph("target")
    comment = document.add_comment(target.runs[0], text="one\ntwo\rthree", author="A")
    assert comment.text == "one\ntwo\nthree", "python-docx's: paragraphs a \\n apart, the \\r a w:br"
    assert len(comment.paragraphs) == 2
    with pytest.raises(ValueError):
        document.add_comment(target.runs[0], text="a\u0002b", author="A")


def _run_of(document, text):
    paragraph = document.add_paragraph()
    paragraph.add_run(text)
    return paragraph


def _set_run_text(document, text):
    paragraph = document.add_paragraph("x")
    paragraph.runs[0].text = text
    return paragraph


def _set_paragraph_text(document, text):
    paragraph = document.add_paragraph("x")
    paragraph.text = text
    return paragraph


# ---------------------------------------------------------------------------
# the readers that do not move
# ---------------------------------------------------------------------------


def test_the_readers_that_keep_docx4js_characters():
    package = fixture("48-text.docx")
    paragraph = next(p for p in package.body.paragraphs if p.text.startswith("Case 3:"))
    assert text_of(paragraph.element).endswith("LINEA\nLINEB"), "traversal.text_of is docx4j's TextUtils"
    assert text_of_view(paragraph.element, characters="docx4j").endswith("LINEA\nLINEB")
    from docx4j_py.docx import Document

    facade = Document(str(OFFICE_JS / "48-text.docx"))
    assert any(p.text.endswith("LINEA\nLINEB") for p in facade.paragraphs), "python-docx's Run.text (CR-011 Q5)"
    hyphens = next(p for p in facade.paragraphs if "HYPHA" in p.text)
    assert hyphens.text.endswith("HYPHA‑HYPHB­HYPHC")
    assert [element_name(item) for item in paragraph.element.content][-1]
