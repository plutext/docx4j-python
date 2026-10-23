"""Word's own writers: CR-013 section 11, Phase A's amendment (decided 2026-10-23).

    .venv-fork/bin/python -m pytest tests/content/test_word_writers.py

docx4j-core-ts's Word checks 55 (``a923aeb``) and 55b (``1e61bb1``) measured what
Word 365's writers do with the characters Office JS's readers give; core-ts
decided its CR-002 section 60 on them (sections 60.4 to 60.7), and the user here
("let's fix core-ts's E1 and E3 now"; for E3, "Drop, as Word does"):

* **E1.** Every writer through ``Paragraph.splice`` --- ``insert_text`` at any
  location, a range's, the body's, a block control's, and ``replace_text`` ---
  makes a new paragraph of each ``\\n`` and each ``\\r``, ``\\r\\n`` one. Both
  halves keep the paragraph's properties; tracked, the new mark is marked
  inserted; a split inside a hyperlink, a revision or a smart tag cuts it in
  two; one inside a run-level content control, a field, an equation or a kept
  ``mc:AlternateContent`` is refused before anything changes.
* **E2.** ``insert_paragraph`` and ``Paragraph.text = ...`` keep one paragraph:
  ``\\r\\n`` is one ``w:br``, where Phase A wrote a ``w:cr`` and a ``w:br``.
* **E3.** U+0002 and the other characters XML 1.0 excludes are dropped at the
  write, with no error.
* **From core-ts CR-002 section 61** (``80a933d``), as its Word check 56
  widened it (``38dcae9``, H1 and H4): inside a bound block control --- no type,
  ``w:text`` or ``w:richText``, all of which Word reads as plain text --- a
  ``\\n`` or ``\\r`` makes no paragraph: a ``w:br`` in a multi-line ``w:text``
  control, nothing in any other.

The oracles are Word's answers, written here as data and cited by commit and
case: core-ts's ``.docx`` and JSON files are not copied (CR-013 Q8's grant names
checks 48 to 53 only). Each case's file is check 55's region (``BEGIN<n>``, the
case's paragraph, ``END<n>``), built here.
"""

from __future__ import annotations

import re

import pytest
from lxml import etree

from docx4j_py import create_package
from docx4j_py.model.content.errors import AddressError, ContentError, SpanError
from docx4j_py.model.revisions import accept_all, reject_all
from docx4j_py.wml import to_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def xml_of(package) -> str:
    return to_xml(package.main_document_part.contents)


def shape(element) -> str:
    """A paragraph as tokens: its mark's insertion, its text merged across runs, each element by name, ``w:ins`` brackets.

    The form check 55's OOXML is compared in, so that run boundaries, run
    properties, ids and dates (which differ between Word and this package and
    mean nothing to the reader) do not count.
    """
    root = etree.fromstring(to_xml(element).encode())
    tokens: list[str] = []
    if root.find(f"{{{W}}}pPr/{{{W}}}rPr/{{{W}}}ins") is not None:
        tokens.append("¶+")

    def walk(node) -> None:
        for child in node:
            name = etree.QName(child).localname
            if name == "pPr":
                continue
            if name == "ins":
                tokens.append("+[")
                walk(child)
                tokens.append("]")
                continue
            if name == "r":
                for item in child:
                    local = etree.QName(item).localname
                    if local == "rPr":
                        continue
                    if local == "t":
                        if tokens and tokens[-1].startswith('"'):
                            tokens[-1] = tokens[-1][:-1] + (item.text or "") + '"'
                        else:
                            tokens.append('"' + (item.text or "") + '"')
                    elif local == "br":
                        kind = item.get(f"{{{W}}}type")
                        tokens.append(f"<br {kind}/>" if kind else "<br/>")
                    else:
                        tokens.append(f"<{local}/>")
                continue
            walk(child)

    walk(root)
    return " ".join(tokens)


def region(n: int, *, tracked: bool = False):
    """Check 55's region for case `n`: ``BEGIN<n>``, ``TARGET<n>``, ``END<n>``."""
    package = create_package()
    for text in (f"BEGIN{n}", f"TARGET{n}", f"END{n}"):
        package.body.insert_paragraph(text)
    if tracked:
        package.change_tracking_mode = "TrackAll"
    return package, package.body.paragraphs[1]


def between(package, n: int) -> list:
    """The paragraphs between case `n`'s ``BEGIN`` and ``END``."""
    paragraphs = package.body.paragraphs
    first = next(i for i, p in enumerate(paragraphs) if p.text == f"BEGIN{n}")
    last = next(i for i, p in enumerate(paragraphs) if p.text == f"END{n}")
    return paragraphs[first + 1 : last]


def chars(spec: str) -> str:
    return "".join(chr(int(code.removeprefix("U+"), base=16)) for code in spec.split(","))


def strict_reload(package):
    """Save, scan the saved document part for raw control characters, and load again strictly."""
    from docx4j_py import load
    from docx4j_py.openpackaging.load import LoadOptions

    data = package.save()
    again = load(data, options=LoadOptions(strict=True))
    assert not again.main_document_part.skipped
    assert not re.search("[\x00-\x08\x0b\x0c\x0e-\x1f]", xml_of(again)), "no raw control character saved"
    return again


# ---------------------------------------------------------------------------
# check 55 (core-ts a923aeb): insertText 'End', insertParagraph 'After', insertComment, comment.content,
# insertText 'End' tracked. Word's answers in its reopened save (55-characters-word365-reopened.json):
# (case, operation, tracked, characters, the region's paragraphs, their markup's shape, insertText's range text in
# Word's session).
# ---------------------------------------------------------------------------

CHECK_55 = [
    (1, "text-end", False, "U+0009", ["TARGET1x\ty"], ['"TARGET1x" <tab/> "y"'], "x\ty"),
    (2, "text-end", False, "U+000A", ["TARGET2x", "y"], ['"TARGET2x"', '"y"'], "x\ry"),
    (3, "text-end", False, "U+000D", ["TARGET3x", "y"], ['"TARGET3x"', '"y"'], "x\ry"),
    (4, "text-end", False, "U+000D,U+000A", ["TARGET4x", "y"], ['"TARGET4x"', '"y"'], "x\ry"),
    (5, "text-end", False, "U+000B", ["TARGET5x\x0by"], ['"TARGET5x" <br/> "y"'], "x\x0by"),
    (6, "text-end", False, "U+000C", ["TARGET6x\x0cy"], ['"TARGET6x" <br page/> "y"'], "x\x0cy"),
    (7, "text-end", False, "U+000E", ["TARGET7x\x0ey"], ['"TARGET7x" <br column/> "y"'], "x\x0ey"),
    (8, "text-end", False, "U+001E", ["TARGET8x\x1ey"], ['"TARGET8x" <noBreakHyphen/> "y"'], "x\x1ey"),
    (9, "text-end", False, "U+001F", ["TARGET9x\x1fy"], ['"TARGET9x" <softHyphen/> "y"'], "x\x1fy"),
    (10, "text-end", False, "U+0002", ["TARGET10xy"], ['"TARGET10xy"'], "x\x02y"),
    (11, "text-end", False, "U+0001", ["TARGET11xy"], ['"TARGET11xy"'], "x\x01y"),
    (12, "text-end", False, "U+0007", ["TARGET12xy"], ['"TARGET12xy"'], "xy"),
    (13, "text-end", False, "U+2011", ["TARGET13x‑y"], ['"TARGET13x‑y"'], "x‑y"),
    (14, "text-end", False, "U+00AD", ["TARGET14x\xady"], ['"TARGET14x\xady"'], "x\xady"),
    (15, "paragraph-after", False, "U+0009", ["TARGET15", "x\ty"], ['"TARGET15"', '"x" <tab/> "y"'], "x\ty"),
    (16, "paragraph-after", False, "U+000A", ["TARGET16", "x\\ny"], ['"TARGET16"', '"x\\ny"'], "x\\ny"),
    (17, "paragraph-after", False, "U+000D", ["TARGET17", "x\\ry"], ['"TARGET17"', '"x\\ry"'], "x\\ry"),
    (18, "paragraph-after", False, "U+000D,U+000A", ["TARGET18", "x\\r\\ny"], ['"TARGET18"', '"x\\r\\ny"'], "x\\r\\ny"),
    (19, "paragraph-after", False, "U+000B", ["TARGET19", "x\x0by"], ['"TARGET19"', '"x" <br/> "y"'], "x\x0by"),
    (20, "paragraph-after", False, "U+000C", ["TARGET20", "x\x0cy"], ['"TARGET20"', '"x" <br page/> "y"'], "x\x0cy"),
    (21, "paragraph-after", False, "U+000E", ["TARGET21", "x\x0ey"], ['"TARGET21"', '"x" <br column/> "y"'], "x\x0ey"),
    (22, "paragraph-after", False, "U+001E", ["TARGET22", "x\x1ey"], ['"TARGET22"', '"x" <noBreakHyphen/> "y"'], "x\x1ey"),
    (23, "paragraph-after", False, "U+001F", ["TARGET23", "x\x1fy"], ['"TARGET23"', '"x" <softHyphen/> "y"'], "x\x1fy"),
    (24, "paragraph-after", False, "U+0002", ["TARGET24", "xy"], ['"TARGET24"', '"xy"'], "x\x02y"),
    (25, "paragraph-after", False, "U+0001", ["TARGET25", "xy"], ['"TARGET25"', '"xy"'], "x\x01y"),
    (26, "paragraph-after", False, "U+0007", ["TARGET26", "xy"], ['"TARGET26"', '"xy"'], "xy"),
    (27, "paragraph-after", False, "U+2011", ["TARGET27", "x‑y"], ['"TARGET27"', '"x‑y"'], "x‑y"),
    (28, "paragraph-after", False, "U+00AD", ["TARGET28", "x\xady"], ['"TARGET28"', '"x\xady"'], "x\xady"),
    (47, "text-end", True, "U+0009", ["TARGET47x\ty"], ['"TARGET47" +[ "x" <tab/> "y" ]'], "x\ty"),
    (48, "text-end", True, "U+000A", ["TARGET48x", "y"], ['¶+ "TARGET48" +[ "x" ]', '+[ "y" ]'], "x\ry"),
    (49, "text-end", True, "U+000D", ["TARGET49x", "y"], ['¶+ "TARGET49" +[ "x" ]', '+[ "y" ]'], "x\ry"),
    (50, "text-end", True, "U+000B", ["TARGET50x\x0by"], ['"TARGET50" +[ "x" <br/> "y" ]'], "x\x0by"),
    (51, "text-end", True, "U+001E", ["TARGET51x\x1ey"], ['"TARGET51" +[ "x" <noBreakHyphen/> "y" ]'], "x\x1ey"),
]

#: Where this package's answer is not Word's, by a recorded departure: case to (paragraphs, shapes, why).
CHECK_55_DEPARTURES = {
    # E2, a recorded departure (core-ts CR-002 section 60.4): insertParagraph keeps one paragraph and the break,
    # where Word writes the escapes as text (a backslash, then n or r)
    16: (["TARGET16", "x\x0by"], ['"TARGET16"', '"x" <br/> "y"'], "E2: \\n a w:br"),
    17: (["TARGET17", "x\ry"], ['"TARGET17"', '"x" <cr/> "y"'], "E2: \\r a w:cr"),
    18: (["TARGET18", "x\x0by"], ['"TARGET18"', '"x" <br/> "y"'], "E2: \\r\\n one w:br"),
}

#: The range ``insert_text`` returns, where it is not Word's (case to (ours, why)).
CHECK_55_RANGE_DEPARTURES = {
    # cases 2 to 4, 48 and 49: the range across the paragraphs the text made is Word's (x\ry) since
    # CR-013 Phase B, where it was the text's part in the last paragraph (y) until then
    # E3: dropped at the write, where Word's session read them until the save (its save holds neither)
    10: ("xy", "dropped at once"),
    11: ("xy", "dropped at once"),
    24: ("xy", "dropped at once"),
    25: ("xy", "dropped at once"),
    # E2 (above): the new paragraph's text
    16: ("x\x0by", "E2"),
    17: ("x\ry", "E2"),
    18: ("x\x0by", "E2"),
}


@pytest.mark.parametrize(
    ("n", "op", "tracked", "spec", "paragraphs", "shapes", "returned"),
    CHECK_55,
    ids=[f"case{case[0]}-{case[1]}-{case[3]}{'-tracked' if case[2] else ''}" for case in CHECK_55],
)
def test_check_55_word_writers(n, op, tracked, spec, paragraphs, shapes, returned):
    package, target = region(n, tracked=tracked)
    text = "x" + chars(spec) + "y"
    if op == "text-end":
        out = target.insert_text(text, location="End")
        got = out.text
    else:
        got = target.insert_paragraph(text, location="After").text
    if n in CHECK_55_DEPARTURES:
        paragraphs, shapes, _why = CHECK_55_DEPARTURES[n]
    if n in CHECK_55_RANGE_DEPARTURES:
        returned, _why = CHECK_55_RANGE_DEPARTURES[n]
    made = between(package, n)
    assert [p.text for p in made] == paragraphs
    assert [shape(p.element) for p in made] == shapes
    assert got == returned
    again = strict_reload(package)
    assert [p.text for p in between(again, n)] == paragraphs, "it saves, and reads back as written"


#: Check 55's comments (cases 29 to 46): Word's ``comment.content`` after its save.
CHECK_55_COMMENTS = [
    (29, "comment", "U+0009", "x\ty"),
    (30, "comment", "U+000A", "x\ry"),
    (31, "comment", "U+000D", "x\ry"),
    (32, "comment", "U+000D,U+000A", "x\ry"),
    (33, "comment", "U+000B", "x\x0by"),
    (34, "comment", "U+000C", "x\x0cy"),
    (35, "comment", "U+000E", "x\x0ey"),
    (36, "comment", "U+001E", "x\x1ey"),
    (37, "comment", "U+001F", "x\x1fy"),
    (38, "comment", "U+0002", "xy"),
    (39, "comment", "U+0001", "xy"),
    (40, "comment", "U+0007", "xy"),
    (41, "comment", "U+2011", "x‑y"),
    (42, "comment", "U+00AD", "x\xady"),
    (43, "comment-content", "U+000D", "x\ry"),
    (44, "comment-content", "U+000A", "x\ry"),
    (45, "comment-content", "U+000D,U+000A", "x\ry"),
    (46, "comment-content", "U+000B", "x\x0by"),
]


@pytest.mark.parametrize(("n", "op", "spec", "content"), CHECK_55_COMMENTS, ids=[f"case{c[0]}-{c[1]}-{c[2]}" for c in CHECK_55_COMMENTS])
def test_check_55_comments(n, op, spec, content):
    """A comment's ``\\r``, ``\\n`` and ``\\r\\n`` (one) are paragraph marks, as Word's (core-ts's D5 with E1)."""
    package, target = region(n)
    text = "x" + chars(spec) + "y"
    hit = target.search(f"TARGET{n}")[0]
    if op == "comment":
        comment = hit.insert_comment(text)
    else:
        comment = hit.insert_comment(f"COMMENT{n}")
        comment.content = text
    assert comment.content == content
    assert len(comment.paragraphs) == content.count("\r") + 1
    assert [p.text for p in between(package, n)] == [f"TARGET{n}"]
    assert strict_reload(package).body.get_comments()[0].content == content


def test_check_55_tracked_changes_read_as_word_lists_them():
    """Cases 47 to 51's listings: one ``Added`` each, its text in docx4j's characters (CR-011 Q6), as before."""
    expected = {
        47: "x\ty",
        48: "x\ry",  # Word's: x\ry
        49: "x\ry",  # Word's: x\ry
        50: "x\ny",  # Word's: x\u000by; TrackedChange.text keeps docx4j's \n for a break (CR-013 section 4.1)
        51: "x‑y",  # Word's: x\u001ey; docx4j's U+2011 for a non-breaking hyphen (the same)
    }
    for n, spec in ((47, "U+0009"), (48, "U+000A"), (49, "U+000D"), (50, "U+000B"), (51, "U+001E")):
        package, target = region(n, tracked=True)
        target.insert_text("x" + chars(spec) + "y", location="End")
        changes = [(c.type, c.text) for c in package.body.get_tracked_changes()]
        assert changes == [("Added", expected[n])], n


# ---------------------------------------------------------------------------
# check 55b (core-ts 1e61bb1): bound run-level plain-text controls given insertText('x' + c + 'y', 'Replace')
# ---------------------------------------------------------------------------

DATA_NS = "urn:check55b"
NAMESPACES = f"xmlns:ns0='{DATA_NS}'"


def bound_run_control(n: int, *, multi_line: bool, value: str = ""):
    package = create_package()
    kind = '<w:text w:multiLine="1"/>' if multi_line else "<w:text/>"
    package.body.insert_xml(
        f'<w:p xmlns:w="{W}"><w:r><w:t xml:space="preserve">BOUND{n} </w:t></w:r><w:sdt><w:sdtPr>'
        f'<w:tag w:val="c55b-{n}"/><w:id w:val="{55100 + n}"/>{kind}</w:sdtPr><w:sdtContent>'
        f"<w:r><w:t>start{n}</w:t></w:r></w:sdtContent></w:sdt></w:p>"
    )
    part = package.custom_xml_parts.add(f'<check55b xmlns="{DATA_NS}"><value{n}>{value or f"start{n}"}</value{n}></check55b>')
    control = package.body.content_controls[0]
    control.xml_mapping.set_mapping(f"/ns0:check55b[1]/ns0:value{n}[1]", NAMESPACES, part)
    return package, control, part


def node_text(part, n: int) -> str:
    return part.select_single_node(f"/ns0:check55b[1]/ns0:value{n}[1]").text


#: (case, multi-line, characters, Word's control text, Word's control shape, Word's node, ours where it differs).
CHECK_55B = [
    (1, True, "U+0041", "xAy", '"xAy"', "xAy", None),
    (2, True, "U+000B", "x\x0by", '"x" <br/> "y"', "x\r\ny", ("x\x0by", '"x" <br/> "y"', "x\ny", "F1: a node's break \\n")),
    (5, True, "U+001E", "x\x1ey", '"x" <noBreakHyphen/> "y"', "x-y", None),  # F2 since the review: the node's -
    (6, True, "U+0002", "x\x02y", '"xy"', "x y", ("xy", '"xy"', "xy", "E3: dropped at the write, so the node has no space")),
    (7, False, "U+0041", "xAy", '"xAy"', "xAy", None),
    # F3 since the review: the break dropped from the control, as Word drops it; the node follows the control
    (8, False, "U+000B", "xy", '"xy"', "x\r\ny", ("xy", '"xy"', "xy", "F3: the node follows the control")),
]


@pytest.mark.parametrize(("n", "multi", "spec", "text", "markup", "node", "ours"), CHECK_55B, ids=[f"case{c[0]}" for c in CHECK_55B])
def test_check_55b_bound_controls(n, multi, spec, text, markup, node, ours):
    package, control, part = bound_run_control(n, multi_line=multi)
    control.insert_text("x" + chars(spec) + "y", location="Replace")
    if ours is not None:
        text, markup, node, _why = ours
    assert control.text == text
    assert node_text(part, n) == node
    assert shape(control.element) == markup
    strict_reload(package)


@pytest.mark.parametrize(("n", "multi", "spec"), [(3, True, "U+000A"), (4, True, "U+000D"), (9, False, "U+000A")])
def test_check_55b_a_paragraph_break_into_a_run_level_control_is_refused(n, multi, spec):
    """Word moved the control to block level, in a paragraph of its own (cases 3, 4, 9); core-ts's F4 refuses, and so does this."""
    package, control, part = bound_run_control(n, multi_line=multi)
    before = xml_of(package)
    with pytest.raises(SpanError) as raised:
        control.insert_text("x" + chars(spec) + "y", location="Replace")
    assert raised.value.code == "range.splits_control"
    assert xml_of(package) == before
    assert node_text(part, n) == f"start{n}"


def test_check_55b_a_nodes_line_feed_shows_as_a_break():
    """Case 10: Word showed ``line1``, a ``w:br``, ``line2``; docx4j's ``runsForValue`` writes the same here."""
    package, control, _part = bound_run_control(10, multi_line=True, value="line1\nline2")
    package.custom_xml_parts.apply_bindings()
    assert control.text == "line1\x0bline2"


# ---------------------------------------------------------------------------
# E1: a paragraph of each \n and \r, through every writer through the splice
# ---------------------------------------------------------------------------

SPLICE_WRITERS = {
    "Paragraph.insert_text End": lambda p, text: p.insert_text(text, location="End"),
    "Paragraph.insert_text Start": lambda p, text: p.insert_text(text, location="Start"),
    "Paragraph.insert_text Replace": lambda p, text: p.insert_text(text, location="Replace"),
    "Range.insert_text Replace": lambda p, text: p.search("ee")[0].insert_text(text, location="Replace"),
    "Range.insert_text Before": lambda p, text: p.search("ee")[0].insert_text(text, location="Before"),
    "Range.insert_text After": lambda p, text: p.search("ee")[0].insert_text(text, location="After"),
    "Paragraph.replace_text": lambda p, text: p.replace_text("ee", text),
    "Body.replace_text": lambda p, text: p.parent_body.replace_text("ee", text),
    "Body.insert_text End": lambda p, text: p.parent_body.insert_text(text, location="End"),
}

EXPECTED = {
    "Paragraph.insert_text End": ["seedx", "y"],
    "Paragraph.insert_text Start": ["x", "yseed"],
    "Paragraph.insert_text Replace": ["x", "y"],
    "Range.insert_text Replace": ["sx", "yd"],
    "Range.insert_text Before": ["sx", "yeed"],
    "Range.insert_text After": ["seex", "yd"],
    "Paragraph.replace_text": ["sx", "yd"],
    "Body.replace_text": ["sx", "yd"],
    "Body.insert_text End": ["seedx", "y"],
}


def seeded(tracked: bool):
    package = create_package()
    p = package.body.insert_paragraph("seed", style="Heading1")
    p.alignment = "Centered"
    if tracked:
        package.change_tracking_mode = "TrackAll"
    return package, p


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("brk", ["\n", "\r", "\r\n"], ids=["LF", "CR", "CRLF"])
@pytest.mark.parametrize("writer", sorted(SPLICE_WRITERS))
def test_each_splice_writer_makes_a_paragraph_of_each_break(writer, brk, tracked):
    package, p = seeded(tracked)
    SPLICE_WRITERS[writer](p, f"x{brk}y")
    paragraphs = package.body.paragraphs
    assert [q.text for q in paragraphs] == EXPECTED[writer]
    for q in paragraphs:
        assert q.style_id == "Heading1", "both halves keep the paragraph's properties"
        assert q.alignment == "Centered"
    xml = xml_of(package)
    assert "<w:cr/>" not in xml and "<w:br/>" not in xml and "&#13;" not in xml
    if tracked:
        head = paragraphs[0].element
        assert head.p_pr.r_pr is not None and head.p_pr.r_pr.ins is not None, "the new mark is marked inserted"
        assert paragraphs[-1].element.p_pr.r_pr is None or paragraphs[-1].element.p_pr.r_pr.ins is None
        rejected = strict_reload(package)
        reject_all(rejected)
        assert [q.text for q in rejected.body.paragraphs] == ["seed"], "Reject All joins them back"
        accepted = strict_reload(package)
        accept_all(accepted)
        assert [q.text for q in accepted.body.paragraphs] == EXPECTED[writer]
    assert [q.text for q in strict_reload(package).body.paragraphs] == EXPECTED[writer]


def test_the_paragraph_keeps_its_identity_and_the_new_one_comes_in_front():
    """The ``w:p`` keeps the last piece, its mark, its section break and its paraId (core-ts's splitParagraph)."""
    package = create_package()
    p = package.body.insert_paragraph("first")
    p.element.p_pr = None
    second = package.body.insert_paragraph("second")
    para_id = second.para_id
    out = second.insert_text("x\r\ny\rz", location="End")
    assert [q.text for q in package.body.paragraphs] == ["first", "secondx", "y", "z"]
    assert second.text == "z", "the view goes on viewing the paragraph that keeps the mark"
    assert second.para_id == para_id
    assert len({q.para_id for q in package.body.paragraphs}) == 4, "each new paragraph its own paraId"
    # the range returned spans the paragraphs the text made, as Office JS's (CR-013 Phase B); it was the
    # last paragraph's part, (second, 0, 1), until then
    assert (out.paragraph.text, out.start, out.end_paragraph.element is second.element, out.end) == ("secondx", 6, True, 1)
    assert out.text == "x\ry\rz"
    report = package.changes[-1]
    assert report.operation == "insert_text"
    assert len(report.created_para_ids) == 2
    assert report.text_after == "secondx\ry\rz"


def test_consecutive_leading_and_trailing_breaks():
    package = create_package()
    p = package.body.insert_paragraph("ab")
    p.insert_text("\nx\n\ny\n", location="End")
    assert [q.text for q in package.body.paragraphs] == ["ab", "x", "", "y", ""]
    package = create_package()
    p = package.body.insert_paragraph("ab")
    out = p.insert_text("\n", location="Start")
    assert [q.text for q in package.body.paragraphs] == ["", "ab"]
    assert (out.start, out.end) == (0, 0)


def test_a_range_follows_its_text():
    package = create_package()
    p = package.body.insert_paragraph("one two three")
    before = p.search("three")[0]
    before.insert_text("A\nB ", location="Before")
    assert (before.paragraph.text, before.text) == ("B three", "three")
    after = package.body.paragraphs[0].search("one")[0]
    after.insert_text(" C\nD", location="After")
    assert after.text == "one" and after.paragraph.text == "one C"
    replaced = package.body.paragraphs[-1].search("three")[0]
    out = replaced.insert_text("E\nF", location="Replace")
    # across the paragraphs the replacement made since CR-013 Phase B (("F", "F") until then)
    assert (replaced.text, out.text) == ("E\rF", "E\rF")
    assert [q.text for q in package.body.paragraphs] == ["one C", "D two A", "B E", "F"]


def test_replace_text_splits_at_every_match():
    package = create_package()
    package.body.insert_paragraph("a, b, c")
    package.body.insert_paragraph("d, e")
    assert package.body.replace_text(", ", "\r\n") == 3
    assert [q.text for q in package.body.paragraphs] == ["a", "b", "c", "d", "e"]


def test_the_body_and_a_block_controls_insert_text():
    package = create_package()
    package.body.insert_paragraph("old")
    out = package.body.insert_text("x\ny", location="Replace")
    assert [q.text for q in package.body.paragraphs] == ["x", "y"]
    assert out.text == "x\ry", "across the paragraphs it made (CR-013 Phase B; 'y' until then)"
    empty = create_package()
    empty.body.content.clear()
    empty.body.insert_text("p\rq", location="End")
    assert [q.text for q in empty.body.paragraphs] == ["p", "q"]
    package.body.insert_xml(f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="4"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>in</w:t></w:r></w:p></w:sdtContent></w:sdt>')
    control = package.body.content_controls[0]
    control.insert_text("a\nb", location="End")
    assert [q.text for q in control.paragraphs] == ["ina", "b"]
    assert control.text == "ina\rb"
    strict_reload(package)


# ---------------------------------------------------------------------------
# E1: what a split cuts in two, and what it refuses
# ---------------------------------------------------------------------------


def one(xml: str, *, tracked: bool = False):
    package = create_package()
    package.body.insert_xml(f'<w:p xmlns:w="{W}" xmlns:m="{M}" xmlns:mc="{MC}">{xml}</w:p>')
    if tracked:
        package.change_tracking_mode = "TrackAll"
    return package, package.body.paragraphs[0]


LINK = '<w:r><w:t xml:space="preserve">see </w:t></w:r><w:hyperlink w:anchor="here"><w:r><w:t>LINKTEXT</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> end</w:t></w:r>'
OTHERS = '<w:r><w:t xml:space="preserve">a </w:t></w:r><w:ins w:id="7" w:author="Other" w:date="2020-01-01T00:00:00Z"><w:r><w:t>INSERTED</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> b</w:t></w:r>'
SMART = '<w:smartTag w:uri="urn:x" w:element="place"><w:r><w:t>SmartTag</w:t></w:r></w:smartTag>'


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_a_split_inside_a_hyperlink_a_revision_or_a_smart_tag_cuts_it_in_two(tracked):
    package, p = one(LINK, tracked=tracked)
    p.search("LINK")[0].insert_text("x\ny", location="After")
    first, second = package.body.paragraphs
    assert (first.text, second.text) == ("see LINKx", "yTEXT end")
    assert '<w:hyperlink w:anchor="here">' in to_xml(first.element) and '<w:hyperlink w:anchor="here">' in to_xml(second.element)

    package, p = one(OTHERS, tracked=tracked)
    p.search("INS")[0].insert_text("\n", location="After")
    first, second = package.body.paragraphs
    assert (first.text, second.text) == ("a INS", "ERTED b")
    ids = re.findall(r'<w:ins w:id="(\d+)" w:author="Other"', xml_of(package))
    assert len(ids) == 2 and ids[0] == "7" and ids[1] != "7", "the second half under a new id"

    package, p = one(SMART, tracked=tracked)
    p.insert_text("\r", location="Start")
    p.search("Tag")[0].insert_text("\r", location="Before")
    assert [q.text for q in package.body.paragraphs] == ["", "Smart", "Tag"]
    assert xml_of(package).count("<w:smartTag") == 2
    strict_reload(package)


REFUSED = [
    ("a run-level content control",
     '<w:r><w:t xml:space="preserve">AB </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:r><w:t>CD</w:t></w:r></w:sdtContent></w:sdt>',
     lambda p: p.search("C")[0].insert_text("x\ny", location="After"), "range.splits_control"),
    ("text written at a control's end, which goes into it",
     '<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:r><w:t>CD</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> EF</w:t></w:r>',
     lambda p: p.search(" EF")[0].insert_text("x\n", location="Before"), "range.splits_control"),
    ("a complex field's result",
     '<w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>12</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
     lambda p: p.search("1")[0].insert_text("\n", location="After"), "range.splits_field"),
    ("a w:fldSimple",
     '<w:r><w:t xml:space="preserve">A </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>34</w:t></w:r></w:fldSimple>',
     lambda p: p.search("3")[0].insert_text("x\ny", location="After"), "range.splits_field"),
    ("an equation",
     '<w:r><w:t xml:space="preserve">a </w:t></w:r><m:oMath><m:r><m:t>xy</m:t></m:r></m:oMath><w:r><w:t xml:space="preserve"> b</w:t></w:r>',
     lambda p: p.splice(3, 3, "\n"), "range.cuts_equation"),
    ("a kept mc:AlternateContent branch",
     '<w:r><w:t xml:space="preserve">a </w:t></w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>BRANCH</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>BRANCH</w:t></w:r></mc:Fallback></mc:AlternateContent>',
     lambda p: p.splice(4, 4, "x\ny"), "address.read_only"),
]


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize(("label", "xml", "write", "code"), REFUSED, ids=[case[0] for case in REFUSED])
def test_a_split_that_cannot_be_made_is_refused_before_anything_changes(label, xml, write, code, tracked):
    package, p = one(xml, tracked=tracked)
    before = xml_of(package)
    with pytest.raises(ContentError) as raised:
        write(p)
    assert raised.value.code == code, label
    assert isinstance(raised.value, AddressError if code.startswith("address") else SpanError)
    assert xml_of(package) == before, "nothing changed"


def test_beside_a_control_or_a_field_the_split_is_made():
    control = '<w:r><w:t xml:space="preserve">AB </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:r><w:t>CD</w:t></w:r></w:sdtContent></w:sdt>'
    package, p = one(control)
    p.search("AB ")[0].insert_text("x\ny", location="After")
    assert [q.text for q in package.body.paragraphs] == ["AB x", "yCD"]
    field = '<w:r><w:t xml:space="preserve">A </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>34</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> B</w:t></w:r>'
    package, p = one(field)
    p.search("A ")[0].insert_text("x\n", location="After")
    first, second = package.body.paragraphs
    assert (first.text, second.text) == ("A x", "34 B")
    assert "fldSimple" in to_xml(second.element) and "fldSimple" not in to_xml(first.element)
    strict_reload(package)


def test_every_match_is_checked_before_the_first_is_replaced():
    field = '<w:r><w:t xml:space="preserve">x-y </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>a-b</w:t></w:r></w:fldSimple>'
    package, _p = one(field)
    before = xml_of(package)
    with pytest.raises(SpanError):
        package.body.replace_text("-", "\n")
    assert xml_of(package) == before


# ---------------------------------------------------------------------------
# core-ts CR-002 section 61 (80a933d), widened by its check 56 (38dcae9, H1, H4): no paragraph in a bound block
# control, a w:br only in a multi-line w:text one
# ---------------------------------------------------------------------------

DATA_56 = "urn:check56"
FLAT_OPC = (
    '<?xml version="1.0" standalone="yes"?><?mso-application progid="Word.Document"?>'
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"></pkg:package>'
)


def block_control(kind: str, *, bound: bool, value: str = "Note"):
    """A block-level control of `kind` holding `value`, bound (or not) to ``value1`` of a custom XML part holding it."""
    package = create_package()
    package.body.insert_xml(
        f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="9"/>{kind}</w:sdtPr><w:sdtContent>'
        f"<w:p><w:r><w:t>{value}</w:t></w:r></w:p></w:sdtContent></w:sdt>"
    )
    control = package.body.content_controls[0]
    if bound:
        from xml.sax.saxutils import escape

        node = escape(FLAT_OPC) if value == "flat" else value
        part = package.custom_xml_parts.add(f'<check56 xmlns="{DATA_56}"><value1>{node}</value1></check56>')
        control.xml_mapping.set_mapping("/ns0:check56[1]/ns0:value1[1]", f"xmlns:ns0='{DATA_56}'", part)
    return package, control


SECTION_61 = [
    ("w:text multi-line", '<w:text w:multiLine="1"/>', True, ["Notex\x0by\x0bz"]),
    ("w:text single-line", "<w:text/>", True, ["Notexyz"]),
    ("no type", "", True, ["Notexyz"]),
    ("w:richText", "<w:richText/>", True, ["Notexyz"]),
    ("unbound, no type", "", False, ["Notex", "y", "z"]),
    ("unbound, w:richText", "<w:richText/>", False, ["Notex", "y", "z"]),
]


@pytest.mark.parametrize("writer", ["control", "paragraph"])
@pytest.mark.parametrize(("label", "kind", "bound", "paragraphs"), SECTION_61, ids=[case[0] for case in SECTION_61])
def test_a_bound_block_control_holds_one_paragraph(label, kind, bound, paragraphs, writer):
    """Word reads every bound block control as plain text: a ``w:br`` in a multi-line ``w:text`` one, nothing in any other."""
    package, control = block_control(kind, bound=bound)
    if writer == "control":
        control.insert_text("x\ny\r\nz", location="End")
    else:
        control.paragraphs[0].insert_text("x\ny\r\nz", location="End")
    assert [p.text for p in control.paragraphs] == paragraphs
    assert control.type == ("PlainText" if "w:text" in kind else "RichText"), "the reader is CR-014's (H4)"
    strict_reload(package)


def test_a_control_whose_node_holds_flat_opc_is_left_as_it_is():
    """H1: Word's own rich-text binding, its node a flat OPC package, isn't in the rule (not measured): E1 splits."""
    _package, control = block_control("<w:richText/>", bound=True, value="flat")
    control.paragraphs[0].insert_text("x\ny", location="End")
    assert [p.text for p in control.paragraphs] == ["flatx", "y"]


#: Check 56 (core-ts 38dcae9, 56-writes-word365-run.json): insertText('x\ny', 'End') on a bound block control holding
#: start<n>, by kind: (case, kind, the control's text, insertText's range, the node after Word's write).
CHECK_56 = [
    (3, "", "start3xy", "xy", "start3xy"),
    (7, '<w:text w:multiLine="1"/>', "start7x\x0by", "x\x0by", "start7x\r\ny"),
    (11, "<w:text/>", "start11xy", "xy", "start11xy"),
    (15, "<w:richText/>", "start15xy", "xy", "start15xy"),
]

#: Where this package's node is not Word's, and why.
CHECK_56_NODES = {
    7: ("start7x\ny", "F1: a node's break \\n"),
    15: ("start15", "update_from_content_controls skips an explicit w:richText (CR-003 section 4); H4 is CR-014's"),
}


@pytest.mark.parametrize(("n", "kind", "text", "returned", "node"), CHECK_56, ids=[f"case{c[0]}" for c in CHECK_56])
def test_check_56_insert_text_with_a_line_feed(n, kind, text, returned, node):
    package, control = block_control(kind, bound=True, value=f"start{n}")
    out = control.insert_text("x\ny", location="End")
    assert (control.text, out.text, len(control.paragraphs)) == (text, returned, 1)
    package.custom_xml_parts.update_from_content_controls()
    if n in CHECK_56_NODES:
        node, _why = CHECK_56_NODES[n]
    assert control.xml_mapping.custom_xml_node.text == node
    strict_reload(package)


# ---------------------------------------------------------------------------
# E2: the one-paragraph writers, \r\n one w:br
# ---------------------------------------------------------------------------

ONE_PARAGRAPH = {
    "Body.insert_paragraph": lambda pkg, text: pkg.body.insert_paragraph(text),
    "Paragraph.insert_paragraph": lambda pkg, text: pkg.body.paragraphs[0].insert_paragraph(text, location="After"),
    "Range.insert_paragraph": lambda pkg, text: pkg.body.paragraphs[0].get_range().insert_paragraph(text),
    "Paragraph.text": lambda pkg, text: setattr(pkg.body.paragraphs[0], "text", text),
    "TableCell.value": lambda pkg, text: setattr(pkg.body.insert_table(1, 1).cell(0, 0), "value", text),
    "insert_table values": lambda pkg, text: pkg.body.insert_table(1, 1, values=[[text]]),
}


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("writer", sorted(ONE_PARAGRAPH))
def test_the_one_paragraph_writers_write_crlf_as_one_break(writer, tracked):
    package = create_package()
    package.body.insert_paragraph("seed")
    if tracked:
        package.change_tracking_mode = "TrackAll"
    before = len(package.body.paragraphs)
    ONE_PARAGRAPH[writer](package, "a\r\nb\rc\nd")
    xml = xml_of(package)
    assert xml.count("<w:br/>") == 2 and xml.count("<w:cr/>") == 1, "\\r\\n one w:br, \\r a w:cr, \\n a w:br"
    texts = [p.text for p in package.body.iter_paragraphs()]
    assert "a\x0bb\rc\x0bd" in texts
    assert len(package.body.paragraphs) <= before + 1, "one paragraph"
    strict_reload(package)


# ---------------------------------------------------------------------------
# E3: dropped, through every writer
# ---------------------------------------------------------------------------

def _on_the_first_paragraph(write):
    return lambda pkg, text: write(pkg.body.paragraphs[0], text)


EVERY_WRITER = {
    **{name: _on_the_first_paragraph(write) for name, write in SPLICE_WRITERS.items()},
    **ONE_PARAGRAPH,
    "Paragraph.insert_comment": lambda pkg, text: pkg.body.paragraphs[0].insert_comment(text),
    "Comment.content": lambda pkg, text: setattr(pkg.body.paragraphs[0].insert_comment("c"), "content", text),
    "Comment.reply": lambda pkg, text: pkg.body.paragraphs[0].insert_comment("c").reply(text),
    "Body.insert_text Replace": lambda pkg, text: pkg.body.insert_text(text, location="Replace"),
    "Table.add_rows": lambda pkg, text: pkg.body.insert_table(1, 1).add_rows(1, values=[[text]]),
}


@pytest.mark.parametrize("char", ["\x02", "\x01", "\x07", "\x00", "\x1d", "￾", "\ud800"])
@pytest.mark.parametrize("writer", sorted(EVERY_WRITER))
def test_every_writer_drops_what_no_element_carries(writer, char):
    package = create_package()
    package.body.insert_paragraph("seed")
    EVERY_WRITER[writer](package, f"x{char}y")
    for part_xml in (xml_of(package), *(to_xml(c.element) for c in package.body.get_comments())):
        assert char not in part_xml
    saved = strict_reload(package)
    comments = saved.body.get_comments()
    texts = [p.text for p in saved.body.iter_paragraphs()] + [c.content for c in comments]
    texts += [reply.content for comment in comments for reply in comment.replies]
    assert any("xy" in text for text in texts), (writer, texts)


def test_a_notes_reference_written_back_is_lost_silently():
    """E3, as core-ts records it: ``paragraph.text = paragraph.text`` loses a note reference (no text makes a note)."""
    package, p = one('<w:r><w:t>a</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:t>b</w:t></w:r>')
    assert p.text == "a\x02b"
    p.text = p.text
    assert p.text == "ab"
    assert "footnoteReference" not in xml_of(package)


def test_the_facade_keeps_python_docxs_refusal():
    """Section 10.4: python-docx 1.2.0's ``ValueError`` stands in the facade, which is python-docx's semantics."""
    from docx4j_py.docx import Document

    document = Document()
    with pytest.raises(ValueError):
        document.add_paragraph("a\x02b")
    paragraph = document.add_paragraph("a\r\nb")
    assert to_xml(paragraph._p).count("<w:br/>") == 2, "python-docx's \\r\\n two w:br"


# ---------------------------------------------------------------------------
# CR-013 section 11, decided after core-ts's consistency review (6285dfa): 1, text into a hidden paragraph
# ---------------------------------------------------------------------------

HIDDEN = '<w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r>'


def _new_run_order(p) -> list[str]:
    """Each run's text, in order, marked ``(hidden)`` where its own ``w:vanish`` hides it."""
    out = []
    for run in p.runs:
        text = "".join(str(getattr(item, "value", "") or "") for item in run.content)
        hidden = getattr(getattr(run, "r_pr", None), "vanish", None) is not None
        out.append(f"{text}(hidden)" if hidden else text)
    return out


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_text_into_an_all_hidden_paragraph_is_seen_mapped_and_at_the_end(tracked):
    """Measured at e79a8e8: it landed before the hidden run, took its w:vanish and escaped the map."""
    package, p = one(HIDDEN, tracked=tracked)
    assert p.text == ""
    out = p.insert_text("VIS\x0bIBLE", location="End")
    assert p.text == "VIS\x0bIBLE", "seen: no w:vanish copied"
    assert _new_run_order(p) == ["HIDDEN(hidden)", "VISIBLE"], "after the hidden run"
    assert "<w:br/>" in to_xml(p.element), "the map: U+000B a w:br"
    assert out.text == "VIS\x0bIBLE"
    assert strict_reload(package).body.paragraphs[0].text == "VIS\x0bIBLE"


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_the_other_writers_into_a_hidden_paragraph(tracked):
    package, p = one(HIDDEN, tracked=tracked)
    p.insert_text("S", location="Start")
    assert _new_run_order(p) == ["S", "HIDDEN(hidden)"], "Start: before the hidden run"
    package, p = one(HIDDEN, tracked=tracked)
    p.get_range("End").insert_text("R")
    assert _new_run_order(p) == ["HIDDEN(hidden)", "R"], "a collapsed range in a paragraph of no visible text: at the end"
    package, p = one(HIDDEN, tracked=tracked)
    p.get_range("Start").insert_text("B", location="Before")
    assert _new_run_order(p) == ["B", "HIDDEN(hidden)"]
    package, p = one(HIDDEN, tracked=tracked)
    p.insert_text("NEW", location="Replace")
    assert p.text == "NEW"
    if not tracked:
        assert _new_run_order(p) == ["NEW"], "Replace takes the hidden run with the whole content"
    package, p = one(HIDDEN, tracked=tracked)
    p.text = "SET"
    assert p.text == "SET"
    package, p = one(HIDDEN, tracked=tracked)
    package.body.insert_text("BODY", location="End")
    assert package.body.paragraphs[-1].text == "BODY"
    assert p.replace_text("HIDDEN", "x") == 0, "hidden text is not found"
    strict_reload(package)


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_at_the_hidden_edge_of_a_partly_hidden_paragraph(tracked):
    package, p = one('<w:r><w:rPr><w:b/></w:rPr><w:t>AB</w:t></w:r>' + HIDDEN, tracked=tracked)
    p.insert_text("X", location="End")
    assert p.text == "ABX"
    order = _new_run_order(p)
    assert order.index("X") > order.index("HIDDEN(hidden)"), "End is after the trailing hidden run"
    added = next(run for run in p.runs if any(getattr(i, "value", None) == "X" for i in run.content))
    assert added.r_pr is not None and added.r_pr.b is not None and added.r_pr.vanish is None, "the visible text's look"
    package, p = one(HIDDEN + '<w:r><w:t>CD</w:t></w:r>', tracked=tracked)
    p.insert_text("Y", location="Start")
    assert p.text == "YCD"
    assert _new_run_order(p)[0] == "Y", "Start is before the leading hidden run"
    strict_reload(package)


# ---------------------------------------------------------------------------
# 2, the break rule in every w:text control, bound or not; 3, F3
# ---------------------------------------------------------------------------

UNBOUND_TEXT_CONTROLS = [
    ("w:text multi-line", '<w:text w:multiLine="1"/>', "Notex\x0by\x0bz"),
    ("w:text single-line", "<w:text/>", "Notexyz"),
]


@pytest.mark.parametrize(("label", "kind", "text"), UNBOUND_TEXT_CONTROLS, ids=[c[0] for c in UNBOUND_TEXT_CONTROLS])
def test_an_unbound_w_text_block_control_takes_a_break_as_its_kind_does(label, kind, text):
    package, control = block_control(kind, bound=False)
    control.insert_text("x\ny\u000bz", location="End")
    assert [p.text for p in control.paragraphs] == [text]
    p = control.paragraphs[0]
    p.text = "a\rb\r\nc"
    assert p.text == ("a\x0bb\x0bc" if "multiLine" in kind else "abc"), "Paragraph.text = too; \\r a w:br in a control"
    strict_reload(package)


def run_control(kind: str):
    package = create_package()
    package.body.insert_xml(
        f'<w:p xmlns:w="{W}"><w:r><w:t xml:space="preserve">at </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>{kind}'
        "</w:sdtPr><w:sdtContent><w:r><w:t>in</w:t></w:r></w:sdtContent></w:sdt></w:p>"
    )
    return package, package.body.content_controls[0]


@pytest.mark.parametrize(("kind", "text"), [('<w:text w:multiLine="1"/>', "inx\x0by"), ("<w:text/>", "inxy"), ("", "inx\x0by")])
def test_a_run_level_control_drops_a_line_break_where_single_line(kind, text):
    """F3, unbound too: a single-line plain-text control takes no line break, through the control and through a range."""
    package, control = run_control(kind)
    control.insert_text("x\x0by", location="End")
    assert control.text == text
    package, control = run_control(kind)
    package.body.paragraphs[0].search("in")[0].insert_text("x\x0by", location="Replace")
    assert control.text == text[2:]
    with pytest.raises(SpanError) as raised:
        control.insert_text("p\nq", location="End")
    assert raised.value.code == "range.splits_control"
    strict_reload(package)


def test_f3_a_bound_single_line_block_control_drops_a_line_break_and_its_node_follows():
    package, control = block_control("<w:text/>", bound=True)
    control.insert_text("x\u000by", location="End")
    package.custom_xml_parts.update_from_content_controls()
    assert (control.text, control.xml_mapping.custom_xml_node.text) == ("Notexy", "Notexy")


def test_f2_a_non_breaking_hyphen_goes_to_the_node_as_a_hyphen_and_an_optional_one_stays():
    package, control, part = bound_run_control(5, multi_line=True)
    control.insert_text("x\u001ey\u001fz", location="Replace")
    assert node_text(part, 5) == "x-y­z"
    strict_reload(package)


# ---------------------------------------------------------------------------
# 3, G1: get_content_controls(types=...)
# ---------------------------------------------------------------------------


def test_get_content_controls_filters_by_type_as_word_does():
    package = create_package()
    package.body.insert_xml(
        f'<w:p xmlns:w="{W}"><w:r><w:t xml:space="preserve">a </w:t></w:r>'
        '<w:sdt><w:sdtPr><w:id w:val="1"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>PLAIN</w:t></w:r></w:sdtContent></w:sdt>'
        '<w:r><w:t xml:space="preserve"> b </w:t></w:r>'
        '<w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent><w:r><w:t>RICH</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    )
    package.body.insert_xml(
        f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="3"/></w:sdtPr><w:sdtContent>'
        '<w:p><w:sdt><w:sdtPr><w:id w:val="4"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>INNER</w:t></w:r></w:sdtContent></w:sdt></w:p>'
        "</w:sdtContent></w:sdt>"
    )
    body = package.body
    ids = lambda controls: [c.id for c in controls]
    assert ids(body.get_content_controls()) == ids(body.content_controls) == [1, 2, 3, 4]
    assert ids(body.get_content_controls(types=["PlainText"])) == [1, 4]
    assert ids(body.get_content_controls(types=["RichText"])) == [2, 3]
    assert ids(body.get_content_controls(types=[])) == [1, 2, 3, 4], "no types: every one, as Office JS's"
    paragraph = body.paragraphs[0]
    assert ids(paragraph.get_content_controls(types=["PlainText"])) == [1]
    block = body.content_controls[2]
    assert ids(block.get_content_controls(types=["PlainText"])) == [4]
    # a range: the controls it overlaps, not one holding it whole (core-ts's rule, check 53)
    assert ids(paragraph.search("a PLA")[0].get_content_controls()) == [1]
    assert ids(paragraph.search("LAI")[0].get_content_controls()) == [], "held whole by control 1"
    assert ids(paragraph.get_range().get_content_controls(types=["RichText"])) == [2]


# ---------------------------------------------------------------------------
# 4, a deletion takes the pictures in its span (core-ts CR-002 section 48)
# ---------------------------------------------------------------------------


def _png() -> bytes:
    """The PNG out of ``samples/Images.docx``."""
    import zipfile
    from pathlib import Path

    samples = Path(__file__).resolve().parents[2] / "samples"
    with zipfile.ZipFile(samples / "Images.docx") as archive:
        return archive.read(next(n for n in archive.namelist() if n.lower().endswith(".png")))


def picture_paragraph(*, tracked: bool = False):
    """``a``, a picture, ``b``: the picture between the two letters, each in a run of its own."""
    from docx4j_py.child import link_parents
    from docx4j_py.wml import R, t

    package = create_package()
    p = package.body.insert_paragraph("a")
    p.insert_inline_picture(_png(), width=20, location="End")
    p.element.content.append(R(content=[t("b")]))
    link_parents(p.element)
    assert p.text == "ab" and len(p.inline_pictures) == 1
    if tracked:
        package.change_tracking_mode = "TrackAll"
    return package, p


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_a_deletion_takes_the_pictures_in_its_span(tracked):
    for write in (
        lambda p: p.get_range("Whole").delete(),
        lambda p: p.get_range("Content").delete(),
        lambda p: p.insert_text("NEW", location="Replace"),
        lambda p: p.search("a^gb")[0].insert_text("NEW", location="Replace"),
        lambda p: setattr(p, "text", "NEW"),
    ):
        package, p = picture_paragraph(tracked=tracked)
        write(p)
        xml = to_xml(p.element)
        if tracked:
            assert re.search(r"<w:del [^>]*>(?:(?!</w:del>).)*<w:drawing>", xml), "the picture's run inside the w:del"
            accepted = strict_reload(package)
            accept_all(accepted)
            assert accepted.body.paragraphs[0].inline_pictures == []
            rejected = strict_reload(package)
            reject_all(rejected)
            assert len(rejected.body.paragraphs[0].inline_pictures) == 1 and rejected.body.paragraphs[0].text == "ab"
        else:
            assert "<w:drawing>" not in xml
        strict_reload(package)


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_a_span_beside_a_picture_leaves_it(tracked):
    """Only a paragraph's whole range holds the pictures at its edges; any other range, those strictly inside it."""
    for find in ("a", "b"):
        _package, p = picture_paragraph(tracked=tracked)
        p.search(find)[0].delete()
        assert "<w:drawing>" in to_xml(p.element)
        if tracked:
            assert not re.search(r"<w:del [^>]*>(?:(?!</w:del>).)*<w:drawing>", to_xml(p.element))


def test_a_tracked_deletion_of_a_paragraph_holding_only_a_picture_takes_it():
    package = create_package()
    p = package.body.insert_paragraph("")
    p.insert_inline_picture(_png(), width=20)
    package.body.insert_paragraph("after")
    package.change_tracking_mode = "TrackAll"
    p.delete()
    xml = to_xml(p.element)
    assert re.search(r"<w:del [^>]*>(?:(?!</w:del>).)*<w:drawing>", xml) and "<w:rPr><w:del " in xml
    accepted = strict_reload(package)
    accept_all(accepted)
    assert [q.text for q in accepted.body.paragraphs] == ["after"]


def test_a_ranges_xml_trims_the_pictures_outside_it():
    """``Range.get_xml`` trims the copy with the same splice: a picture outside the span goes, one at its edge stays.

    core-ts's section 48 says nothing of ``getOoxml``; this is the deletion's
    rule applied to the copy, and the document is not touched.
    """
    from docx4j_py.child import link_parents
    from docx4j_py.wml import R, t

    package = create_package()
    p = package.body.insert_paragraph("ax")
    p.insert_inline_picture(_png(), width=20, location="End")
    p.element.content.append(R(content=[t("b")]))
    link_parents(p.element)
    assert p.text == "axb"
    assert "drawing" not in p.search("a")[0].get_xml(), "the picture after x is outside a's span"
    assert "drawing" in p.search("x")[0].get_xml(), "at x's end edge: kept"
    assert "drawing" in p.search("b")[0].get_xml(), "at b's start edge: kept"
    assert "drawing" in p.get_range().get_xml()
    assert len(p.inline_pictures) == 1, "the document is not touched"
