"""A paragraph break inside a field's result, and a tracked delete of a paragraph holding an equation: CR-013 section 13.7.

    .venv-fork/bin/python -m pytest tests/content/test_fields_equations.py

The port of docx4j-core-ts's ``test/fields-equations.test.mjs`` at ``906d5e3``
(its CR-002 section 62), test for test, its strings kept. docx4j-core-ts's Word
check 57 measured what Word 365 does: a file for a Script Lab snippet, a region
per case, two of them only opened. The snippet's file and the snippet must
agree, since Word runs one on the other and nothing can be fixed after the run.
core-ts's section 62.2 has Word's run, and D1, D2, F1 and F2, which this package
follows since CR-013 section 13.7 (the user, 2026-10-23: "Follow all four"): a
``w:fldSimple`` made a complex field before a break splits it, and a complex
field's result split (where this package refused a break anywhere in a field);
an equation deleted under tracking in Word's form (where this package refused);
Word's form accepted and read as Word has it --- F1 docx4j's own rule since its
``ebb21c990`` (``Resolver.dropEmptiedEquation``, its CR-043 section 16).

core-ts runs check 57's snippet through its Office JS shim; this package has
none, so ``tests/office_js/checks.py`` ports it (``run_check_57``), with the same
JSON shape and ``vis()`` (CR-013 section 4.9 item 6). Every difference from Word
is asserted as it stands, with its reason.
"""

from __future__ import annotations

import functools
import hashlib
import importlib.util
import io
import json
import re
import zipfile

import pytest
from conftest import ROOT

from docx4j_py import create_package, load
from docx4j_py.model import revisions
from docx4j_py.model.content import TrackedChangeError

OFFICE_JS = ROOT / "tests" / "office_js"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_spec = importlib.util.spec_from_file_location("office_js_checks", OFFICE_JS / "checks.py")
checks = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(checks)

#: check 57's cases, as core-ts's ``build.mjs`` (906d5e3) lists them.
CASES = [
    {"n": n, "op": op, "shape": shape}
    for n, (op, shape) in enumerate(
        [
            ("insertText-after", "complex-field"),
            ("insertText-after", "simple-field"),
            ("insertParagraph-after", "complex-field"),
            ("insertParagraph-after", "simple-field"),
            ("open", "display-equation"),
            ("open", "inline-equation"),
            ("paragraph-delete", "display-equation"),
            ("paragraph-delete", "inline-equation"),
            ("range-delete", "display-equation"),
            ("range-delete", "inline-equation"),
        ],
        start=1,
    )
]


def word_run() -> dict:
    """Word 365's run of check 57 (core-ts CR-002 section 62.2), as its snippet reported it."""
    return json.loads((OFFICE_JS / "57-fields-equations-word365-run.json").read_text(encoding="utf-8"))


def saved(name: str) -> bytes:
    return (OFFICE_JS / name).read_bytes()


def unvis(text: str) -> str:
    """A string the snippet made visible (``vis``), as it was."""
    return re.sub(r"\\u([0-9a-f]{4})", lambda m: chr(int(m.group(1), 16)), text).encode("utf-16", "surrogatepass").decode("utf-16")


def comparable(xml: str) -> str:
    """A region's OOXML as two runs can compare it: no revision ids, authors or dates, no xml:space, no last empty paragraph."""
    return re.sub(r"<w:p/>$", "", re.sub(r' (?:w:id|w:author|w:date|w16du:dateUtc|xml:space)="[^"]*"', "", xml))


def paragraphs_of(xml: str) -> list[str]:
    """A region's OOXML split in paragraphs: its BEGIN paragraph, the case's, ..., its END paragraph."""
    return re.findall(r"<w:p\b(?:(?!<w:p\b)[\s\S])*?</w:p>|<w:p/>", xml)


def document_xml_of(data: bytes) -> str:
    return zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml").decode("utf-8")


def body_paragraphs(data: bytes) -> list[str]:
    """A saved document's body in paragraphs, compared as Word saves: no rsids, ids or dates, a run joined to a like neighbour.

    core-ts's ``bodyParagraphs``. This package's saves name their namespaces as
    docx4j's prefix table has them and write booleans as ``1``, where Word's
    may differ: the declarations are left out, as core-ts's are.
    """
    found = re.search(r"<w:body>([\s\S]*)</w:body>", document_xml_of(data))
    assert found is not None
    body = found.group(1)
    body = re.sub(r"<w:sectPr\b[\s\S]*</w:sectPr>", "", body)
    body = body.replace("<w:lastRenderedPageBreak/>", "")
    body = re.sub(
        r' (?:w:rsid\w*|w14:paraId|w14:textId|w:id|w:author|w:date|w16du:dateUtc|xml:space)="[^"]*"| xmlns:\w+="[^"]*"',
        "",
        body,
    )
    was = None
    while was != body:
        was = body
        body = re.sub(
            r"<w:r>((?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?)<w:t>([^<]*)</w:t></w:r><w:r>\1<w:t>([^<]*)</w:t></w:r>",
            r"<w:r>\1<w:t>\2\3</w:t></w:r>",
            body,
        )
        body = re.sub(
            r"<m:r>((?:<w:del>)?(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?)<m:t>([^<]*)</m:t>((?:</w:del>)?)</m:r>"
            r"<m:r>\1<m:t>([^<]*)</m:t>\3</m:r>",
            r"<m:r>\1<m:t>\2\4</m:t>\3</m:r>",
            body,
        )
    return paragraphs_of(body)


def case_paragraph_of(paragraphs: list[str], n: int) -> str:
    """The paragraph after a case's BEGIN paragraph, of a list of paragraphs as :func:`body_paragraphs` gives them."""
    index = next(i for i, p in enumerate(paragraphs) if f">BEGIN{n} " in p)
    return paragraphs[index + 1]


@functools.cache
def our_run() -> tuple[dict, bytes]:
    """Check 57's snippet in its ``run`` mode, through the port on the file Word read, once for the tests below."""
    package = load(OFFICE_JS / "57-fields-equations.docx")
    result = checks.run_check_57(package, "run")
    return result, package.save()


def test_check_57s_snippet_file_and_snippet_agree_every_case_parses_to_the_case_it_was_built_as_of_its_shape():
    package = load(OFFICE_JS / "57-fields-equations.docx")
    paragraphs = package.body.paragraphs
    assert len(CASES) == 10
    assert [r for p in paragraphs if (r := checks.parse_region_57(p.text)) is not None] == CASES
    from docx4j_py.wml import to_xml

    for case in CASES:
        n, op, shape = case["n"], case["op"], case["shape"]
        at = next(i for i, p in enumerate(paragraphs) if p.text == f"BEGIN{n} {op} {shape}")
        p = paragraphs[at + 1]
        assert paragraphs[at + 2].text == f"END{n}"
        xml = to_xml(p.element)
        if shape == "complex-field":
            assert "w:fldChar" in xml and "w:fldSimple" not in xml, f"case {n}"
        elif shape == "simple-field":
            assert "w:fldSimple" in xml and "w:fldChar" not in xml, f"case {n}"
        elif shape == "display-equation":
            assert "m:oMathPara" in xml, f"case {n}"
        else:
            assert "m:oMath" in xml and "m:oMathPara" not in xml, f"case {n}"
        # the words the snippet searches for, each once in the document
        words = (
            [f"FA{n}"]
            if op.startswith("insert")
            else [f"L{n}", f"R{n}"]
            if op == "range-delete" and shape == "inline-equation"
            else []
        )
        for w in words:
            assert len([q for q in paragraphs if w in re.split(r"\W+", q.text)]) == 1, f"case {n}: {w}"
        assert (op == "open") == ("<w:del " in xml), f"case {n}"


def test_check_57s_run_words_json_is_of_this_files_cases_and_its_saves_are_words_as_copied():
    # core-ts 906d5e3's test/check57/manifest.json
    expected = {
        "57-fields-equations-word365.docx": "e8ac302be7a8a7551da61c3a1306881ca0bb5fba",
        "57-fields-equations-word365-accepted.docx": "ee615f9e9a876a4aeddde7924da8c71ea6ec9d65",
        "57-fields-equations-word365-rejected.docx": "ad0937a076ff013b3e6e51c556899f465ebadef0",
    }
    for name, digest in expected.items():
        assert hashlib.sha1(saved(name)).hexdigest() == digest, name
    word = word_run()
    assert word["host"]["version"] == "16.0.20326.20232"
    assert [{"n": c["n"], "op": c["op"], "shape": c["shape"]} for c in word["cases"]] == CASES


def test_cr_002_section_62_2_check_57s_snippet_through_the_port_every_case_goes_ahead_and_only_range_fields_is_missing():
    ours, data = our_run()
    assert len(ours["cases"]) == 10
    assert [c for c in ours["cases"] if "error" in c] == []
    errors = {
        value
        for c in ours["cases"]
        for value in c["region"].values()
        if isinstance(value, str) and value.startswith("error:")
    }
    # Range.fields (Office JS's Word.Field) is not in this package, as it is not in core-ts's shim
    assert errors == {"error: NotImplemented: Range.fields"}
    load(io.BytesIO(data))


def _renumbered(xml: str, origin: int, to: int) -> str:
    """A region's words of case `origin` (keep1a, FA1, END1) as case `to`'s."""
    return re.sub(rf"(?<=[A-Za-z]){origin}(?!\d)", str(to), xml)


def test_cr_002_section_62_2_check_57s_snippet_through_the_port_gives_word_365s_run_case_for_case():
    result, _data = our_run()
    ours = {c["n"]: c for c in result["cases"]}
    word = {c["n"]: c for c in word_run()["cases"]}

    def case_paragraph(case: dict) -> str:
        return paragraphs_of(comparable(case["region"]["ooxml"]))[1]

    for n in range(1, 11):
        assert ours[n].get("result") == word[n].get("result"), f"case {n}"
    # B2: a break inside a field's result splits it, the field running across the paragraphs; a w:fldSimple made a
    # complex field first (D1)
    for n in (1, 2, 3, 4):
        for key in ("text", "paragraphs", "trackedChanges"):
            assert ours[n]["region"][key] == word[n]["region"][key], f"case {n}: {key}"
        assert comparable(ours[n]["region"]["ooxml"]) == comparable(word[n]["region"]["ooxml"]), f"case {n}"
    assert word[2]["region"]["fields"] == [{"code": ' QUOTE "FA2 FB2" ', "result": "FA2x\\u000dy FB2"}]
    # B11: a tracked Paragraph.delete and a tracked range delete taking an equation write Word's form (D2): a w:del
    # in each m:r and in the structure's m:ctrlPr, the runs either side in w:dels of their own
    for n in (7, 8, 9, 10):
        assert comparable(ours[n]["region"]["ooxml"]) == comparable(word[n]["region"]["ooxml"]), f"case {n}"
    # cases 5 and 6 hold the one w:del this package's tracked delete wrote before (core-ts's; this package refused
    # one until CR-013 section 13.7); Word read them as what it writes now
    for n, as_ in ((5, 7), (6, 8)):
        assert _renumbered(case_paragraph(word[n]), n, as_) == case_paragraph(ours[as_]), f"case {n}"
        assert re.search(r"<w:del><(?:m:oMathPara|w:r)>", case_paragraph(ours[n])), f"case {n}"
    # one change a case, as Word lists them. Word's text under tracking is All Markup's, the deleted text in it
    # (core-ts CR-014 section 6.2), where this package's is Simple Markup's; a change's text is docx4j's reader's,
    # the equation's m:t, where Word's is the equation as it draws it
    for n in range(5, 11):
        assert len(ours[n]["region"]["paragraphs"]) == len(word[n]["region"]["paragraphs"]), f"case {n}"
        assert [c["type"] for c in ours[n]["region"]["trackedChanges"]] == [
            c["type"] for c in word[n]["region"]["trackedChanges"]
        ], f"case {n}"
    assert [ours[n]["region"]["paragraphs"][1] for n in range(5, 11)] == ["", "", "", "", "", "keep10a  keep10b"]
    # docx4j's reader (model/revisions/texts.py) reads the equation's m:t in the one-w:del form too, where core-ts's
    # reads none of it there ('\u000d', 'keep6a L6  R6 keep6b\u000d' in core-ts's assertion)
    assert [ours[n]["region"]["trackedChanges"][0]["text"] for n in range(5, 11)] == [
        "E=mc2\\u000d",
        "keep6a L6 E=mc2 R6 keep6b\\u000d",
        "E=mc2\\u000d",
        "keep8a L8 E=mc2 R8 keep8b\\u000d",
        "E=mc2",
        "L10 E=mc2 R10",
    ]
    assert word[10]["region"]["trackedChanges"][0]["text"] == "L10 \\ud835\\udc38=\\ud835\\udc5a,\\ud835\\udc50-2. R10"


def test_cr_002_section_62_2_check_57s_run_here_saved_accepted_and_rejected_gives_words_saves_region_for_region():
    _result, run = our_run()

    def regions(paragraphs: list[str]) -> list[str]:
        return paragraphs[1:-1]

    # the save after the run: every case's paragraphs Word's, but cases 5 and 6, which keep the file's form
    ours = body_paragraphs(run)
    word = body_paragraphs(saved("57-fields-equations-word365.docx"))
    assert len(ours) == len(word)
    differ = [p for p, theirs in zip(regions(ours), regions(word), strict=True) if p != theirs]
    assert differ == [case_paragraph_of(ours, 5), case_paragraph_of(ours, 6)]
    for verb, name in ((revisions.accept_all, "accepted"), (revisions.reject_all, "rejected")):
        resolved = load(io.BytesIO(run))
        verb(resolved)
        assert body_paragraphs(resolved.save()) == body_paragraphs(saved(f"57-fields-equations-word365-{name}.docx")), name


def test_cr_002_section_62_2_words_save_after_check_57s_run_accepted_and_rejected_here_gives_words_accept_all_and_reject_all_f1():
    run = saved("57-fields-equations-word365.docx")
    for verb, name in ((revisions.accept_all, "accepted"), (revisions.reject_all, "rejected")):
        package = load(io.BytesIO(run))
        verb(package)
        assert body_paragraphs(package.save()) == body_paragraphs(saved(f"57-fields-equations-word365-{name}.docx")), name
    # accepted, an equation deleted whole goes with its m:oMath and m:oMathPara (docx4j ebb21c990's
    # dropEmptiedEquation), and its paragraph where the mark went
    accepted = body_paragraphs(saved("57-fields-equations-word365-accepted.docx"))
    assert case_paragraph_of(accepted, 9) == "<w:p/>"
    assert all("oMath" not in p for p in accepted)


def test_cr_002_section_62_2_this_packages_readers_on_words_saves_of_check_57_f2():
    word = {c["n"]: c for c in word_run()["cases"]}

    def case_paragraphs(data: bytes) -> dict[int, str]:
        paragraphs = load(io.BytesIO(data)).body.paragraphs
        out = {}
        for case in CASES:
            index = next(i for i, p in enumerate(paragraphs) if p.text.startswith(f"BEGIN{case['n']} "))
            out[case["n"]] = paragraphs[index + 1].text
        return out

    # rejected, each equation reads as Word read it in the run, where All Markup showed the deleted text
    rejected = case_paragraphs(saved("57-fields-equations-word365-rejected.docx"))
    for n in (5, 6, 7, 8, 9, 10):
        assert rejected[n] == unvis(word[n]["region"]["paragraphs"][1]), f"case {n}"
    # Word's form of an equation deleted whole reads as nothing in Simple Markup's rule, its structure with its runs,
    # as the one w:del form does
    run = case_paragraphs(saved("57-fields-equations-word365.docx"))
    assert [run[n] for n in range(5, 11)] == ["", "", "", "", "", "keep10a  keep10b"]
    ours = case_paragraphs(saved("57-fields-equations.docx"))
    assert [ours[5], ours[6]] == ["", ""]
    # the listing: one change a case, as Word's run lists them, its text docx4j's reader's
    changes = revisions.tracked_changes(load(OFFICE_JS / "57-fields-equations-word365.docx"))
    assert [(c.type, c.text, c.author) for c in changes] == [
        ("Deleted", "E=mc2\r", "docx4j"),
        ("Deleted", "keep6a L6 E=mc2 R6 keep6b\r", "docx4j"),
        ("Deleted", "E=mc2\r", "Author A"),
        ("Deleted", "keep8a L8 E=mc2 R8 keep8b\r", "Author A"),
        ("Deleted", "E=mc2", "Author A"),
        ("Deleted", "L10 E=mc2 R10", "Author A"),
    ]
    assert [len(word[n]["region"]["trackedChanges"]) for n in range(5, 11)] == [1, 1, 1, 1, 1, 1]


def body_of(*paragraphs: str):
    """A created package of these paragraphs (``documentOf``'s), its body."""
    package = create_package()
    for paragraph in paragraphs:
        package.body.insert_xml(paragraph, location="End")
    return package, package.body


def equation_or_field_paragraphs(xml: str) -> list[str]:
    """The paragraphs of a document part that hold something, without namespace declarations or paraIds."""
    found = re.findall(r"<w:p\b(?:(?!<w:p\b)[\s\S])*?</w:p>", xml)
    return [re.sub(r' xmlns:\w+="[^"]*"| w14:\w+="[^"]*"', "", p) for p in found if "<w:t" in p or "m:oMath" in p]


def test_cr_002_section_62_2_d1_a_fldsimple_a_break_splits_is_made_a_complex_field_its_lock_dirty_flag_and_data_on_the_begin_its_runs_look_on_the_field_characters():
    bold = "<w:rPr><w:b/></w:rPr>"
    field = (
        f'<w:p xmlns:w="{W}"><w:r><w:t xml:space="preserve">a </w:t></w:r>'
        f'<w:fldSimple w:instr="DATE" w:dirty="true" w:fldLock="true"><w:r>{bold}<w:t>one two</w:t></w:r></w:fldSimple>'
        '<w:r><w:t xml:space="preserve"> b</w:t></w:r></w:p>'
    )
    for write in (
        lambda p: p.get_range().search("one")[0].insert_paragraph("NEW", location="After"),
        lambda p: p.get_range().search("one")[0].insert_text("\n", location="After"),
    ):
        package, body = body_of(field)
        p = next(q for q in body.paragraphs if "one" in q.text)
        write(p)
        paragraphs = equation_or_field_paragraphs(document_xml_of(package.save()))
        # the first half's new mark aside (in the formatting of the run before the split, both ways here); this
        # package writes booleans 1 (CLAUDE.md), where core-ts's "true"
        assert re.sub(r"<w:pPr>.*?</w:pPr>", "", paragraphs[0]) == (
            f'<w:p><w:r><w:t xml:space="preserve">a </w:t></w:r><w:r>{bold}<w:fldChar w:fldCharType="begin" w:fldLock="1" w:dirty="1"/></w:r>'
            f'<w:r>{bold}<w:instrText>DATE</w:instrText></w:r><w:r>{bold}<w:fldChar w:fldCharType="separate"/></w:r><w:r>{bold}<w:t>one</w:t></w:r></w:p>'
        )
        assert paragraphs[-1] == (
            f'<w:p><w:r>{bold}<w:t xml:space="preserve"> two</w:t></w:r><w:r>{bold}<w:fldChar w:fldCharType="end"/></w:r>'
            '<w:r><w:t xml:space="preserve"> b</w:t></w:r></w:p>'
        )
        assert "fldSimple" not in document_xml_of(package.save())


_SOURCE_10 = (
    f'<w:p xmlns:w="{W}" xmlns:m="{M}"><w:r><w:t xml:space="preserve">keep10a L10 </w:t></w:r>'
    '<m:oMath><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>E=m</m:t></m:r>'
    '<m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr></m:sSupPr>'
    '<m:e><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>c</m:t></m:r></m:e>'
    '<m:sup><m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>'
    '<w:r><w:t xml:space="preserve"> R10 keep10b</w:t></w:r></w:p>'
)


def test_cr_002_section_62_2_d2_a_tracked_delete_of_an_equation_holding_a_revision_of_its_own_puts_it_whole_in_one_wdel_a_tracked_range_replacement_goes_ahead_a_paragraph_one_is_refused():
    revised = (
        f'<w:p xmlns:w="{W}" xmlns:m="{M}"><w:r><w:t xml:space="preserve">x </w:t></w:r>'
        '<m:oMath><m:r><w:ins w:id="90" w:author="Other" w:date="2026-10-01T00:00:00Z"><m:t>q</m:t></w:ins></m:r></m:oMath>'
        '<w:r><w:t xml:space="preserve"> y</w:t></w:r></w:p>'
    )
    package, body = body_of(revised)
    package.change_tracking_mode = "TrackAll"
    next(p for p in body.paragraphs if p.text.startswith("x")).delete()
    assert re.search(
        r'<w:del [^>]*><w:r><w:delText xml:space="preserve">x </w:delText></w:r><m:oMath>',
        equation_or_field_paragraphs(document_xml_of(package.save()))[0],
    )
    package, body = body_of(_SOURCE_10)
    package.change_tracking_mode = "TrackAll"
    p = next(q for q in body.paragraphs if q.text.startswith("keep10a"))
    before = document_xml_of(package.save())
    # Paragraph.insert_text's Replace: where Word puts the text beside a deleted equation isn't measured, refused
    with pytest.raises(TrackedChangeError) as refused:
        p.insert_text("new", location="Replace")
    assert refused.value.code == "tracking.equation" and "tracked replacement of an equation" in str(refused.value)
    assert document_xml_of(package.save()) == before
    # a Range's puts its text before the deletion, as Word's does (check 52 cases 39 to 48), the equation deleted in
    # Word's form
    p.get_range().search("L10")[0].expand_to(p.get_range().search("R10")[0]).insert_text("new", location="Replace")
    assert p.text == "keep10a new keep10b"
    assert re.search(
        r'<w:ins [^>]*><w:r><w:t>new</w:t></w:r></w:ins><w:del [^>]*><w:r><w:delText xml:space="preserve">L10 </w:delText></w:r></w:del>'
        r"<m:oMath><m:r><w:del ",
        equation_or_field_paragraphs(document_xml_of(package.save()))[0],
    )
