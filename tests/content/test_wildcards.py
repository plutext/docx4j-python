"""Word's wildcards across a paragraph's mark, and no hit that cuts a grapheme cluster: CR-016 Phase C.

    .venv-fork/bin/python -m pytest tests/content/test_wildcards.py

CR-016 section 4.7, the port of docx4j-core-ts ``3ba1e68`` (its CR-002 section
32's K7, held to Word 365 by check 59) and ``ebb218e`` (section 32 item 1, the
grapheme filter): ``^13`` is a mark and the pattern ``re.DOTALL``; a range
across paragraphs, or holding its mark, searched across its paragraphs' marks;
``Paragraph.search`` over its ``"Whole"`` range, ``Body.search`` (a control's, a
cell's body's) over its ``"Content"`` range; ``replace_text`` refusing a hit
that holds a mark, before anything changes (Q9); and every search, ``find`` and
``replace_text`` dropping a hit whose start or end cuts a grapheme cluster
(Q8), Latin text untouched. Check 59 itself, and check 58's case 13, are
``test_readers.py``'s. core-ts's ``content.test.mjs`` test "a search never
returns a hit that splits a character" is ported here, its strings kept, with
this package's answers where Python's code points differ (CR-016 section 4.9).
"""

from __future__ import annotations

import re

import pytest

from docx4j_py import create_package
from docx4j_py.model.content import SpanError, text_model
from docx4j_py.model.content.text_model import find_all, search_pattern

WILD = {"match_wildcards": True}
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def texts(hits):
    return [hit.text for hit in hits]


def two_paragraphs(*texts_: str):
    """A package whose body holds one paragraph of each text, in order."""
    package = create_package()
    for text in texts_:
        package.body.insert_paragraph(text, location="End")
    return package


# ---------------------------------------------------------------------------
# core-ts content.test.mjs at 3ba1e68 (CR-002 section 32 item 1, K4, K7)
# ---------------------------------------------------------------------------


def test_a_search_never_returns_a_hit_that_splits_a_character():
    package = create_package()
    emoji = package.body.insert_paragraph("a\U0001f600b", location="End")
    # core-ts: 'the emoji is two UTF-16 units'; here a code point (CR-003 decided question 3)
    assert len(emoji.text) == 3
    # core-ts: "each half of the pair matched the class; neither is a hit; the paragraph's mark is, as Word's (check 58
    # case 13, K7)" --- here the class matches the emoji whole, a code point, and keeps it (CR-016 section 4.9)
    assert texts(emoji.search("[!a-z ]", **WILD)) == ["\U0001f600", "\r"]
    # ? is a code point, the pair whole, as Word's (check 58 case 13; CR-002 section 32, K4), "and the paragraph's
    # mark, as Word's (K7)"
    assert texts(emoji.search("?", **WILD)) == ["a", "\U0001f600", "b", "\r"]
    assert texts(emoji.search("a?b", **WILD)) == ["a\U0001f600b"]
    # "and a replace leaves the emoji whole": core-ts's replaceText finds nothing (0); here the emoji, whole
    assert package.body.replace_text("[!a-z ]", "X", **WILD) == 1
    assert emoji.text == "aXb"

    # a letter and the fatha on it (U+064E) are one cluster: core-ts's \p{L}+ (matchRegExp, which this package has
    # not) stops before the mark; the wildcards' ? and class read the same text
    arabic = package.body.insert_paragraph("\u0628\u064e \u062a", location="End")
    content = arabic.get_range("Content")
    assert texts(content.search("?", **WILD)) == [" ", "\u062a"], (
        "the letter without its mark is not a hit; the bare letter after the space is"
    )
    assert texts(content.search("??", **WILD)) == ["\u0628\u064e", " \u062a"], "a match of the whole cluster is"
    assert texts(content.search("[!a-z ]@", **WILD)) == ["\u0628\u064e", "\u062a"]

    # findAll, the exported span form, drops the same hits; Latin text is untouched by all this
    assert find_all("a\U0001f600b", search_pattern("[!a-z ]", match_wildcards=True)) == [(1, 2)]  # core-ts []
    assert find_all("b\u064e", search_pattern("?", match_wildcards=True)) == []
    assert find_all("cat hat", search_pattern("?at", match_wildcards=True)) == [(0, 3), (4, 7)]


def test_k4_question_mark_takes_an_emoji_whole_x_y_over_x_emoji_y():
    """CR-016 section 4.7: K4 had nothing to port, Python's ``?`` a code point already; this holds it to Word's (check 58 case 13)."""
    package = two_paragraphs("x\U0001f600y")
    paragraph = package.body.paragraphs[0]
    assert texts(paragraph.search("x?y", **WILD)) == ["x\U0001f600y"]
    assert texts(paragraph.get_range("Content").search("?", **WILD)) == ["x", "\U0001f600", "y"]


# ---------------------------------------------------------------------------
# the pattern: ^13 and re.DOTALL
# ---------------------------------------------------------------------------


def test_13_is_a_paragraph_mark_and_the_wildcard_pattern_is_compiled_dotall():
    pattern = search_pattern("Q^13Z", match_wildcards=True)
    assert pattern.fullmatch("Q\rZ") and not pattern.fullmatch("Q^13Z")
    # DOTALL, so ? and * cross any character: \r without it too, \n only with it (CR-016 section 4.7, pinned)
    for query in ("?", "Q?Z", "Q*Z"):
        assert search_pattern(query, match_wildcards=True).flags & re.DOTALL, query
    assert search_pattern("Q?Z", match_wildcards=True).fullmatch("Q\nZ")
    assert search_pattern("Q*Z", match_wildcards=True).fullmatch("Q\r\nZ")
    # a plain query is not a wildcard one: ^13 its three characters, no DOTALL
    assert not search_pattern("^13").flags & re.DOTALL
    assert search_pattern("^13").fullmatch("^13")
    # wildcards still read ^p as the two characters (core-ts ranges.test.mjs), and an escaped caret is a caret
    assert search_pattern("^p", match_wildcards=True).fullmatch("^p")
    assert search_pattern("\\^13", match_wildcards=True).fullmatch("^13")


def test_13_over_text_that_spells_it_finds_the_marks_not_the_characters():
    package = two_paragraphs("x^13y", "z")
    assert texts(package.body.search("^13", **WILD)) == ["\r"]
    assert texts(package.body.search("^13")) == ["^13"]


# ---------------------------------------------------------------------------
# the scopes (CR-016 section 4.7; core-ts Body.mts, Paragraph.mts, Range.mts at 3ba1e68)
# ---------------------------------------------------------------------------


def test_each_scope_reads_the_marks_word_reads_a_paragraph_its_whole_range_a_body_its_content_range():
    package = two_paragraphs("Q", "Z")
    body = package.body
    q, z = body.paragraphs
    # a paragraph: its Whole range, its own mark included, the last paragraph's too
    assert texts(q.search("?", **WILD)) == ["Q", "\r"]
    assert texts(z.search("?", **WILD)) == ["Z", "\r"]
    # a body: its Content range, without the last mark (check 59 case 5's control)
    assert texts(body.search("?", **WILD)) == ["Q", "\r", "Z"]
    assert texts(body.search("[!A-Z]", **WILD)) == ["\r"]
    for query in ("Q?Z", "Q^13Z", "Q*Z"):
        assert texts(body.search(query, **WILD)) == ["Q\rZ"], query
    # its Whole range holds the last mark
    assert texts(body.get_range("Whole").search("?", **WILD)) == ["Q", "\r", "Z", "\r"]
    # a range within one paragraph holding no mark reads none; one across paragraphs the marks between
    assert texts(q.get_range("Content").search("?", **WILD)) == ["Q"]
    across = q.get_range("Content").expand_to(z.get_range("Content"))
    assert texts(across.search("?", **WILD)) == ["Q", "\r", "Z"]
    # the limit, as every search's
    assert texts(body.search("?", limit=2, **WILD)) == ["Q", "\r"]


def test_a_hit_holding_a_mark_is_a_range_that_holds_it_and_one_across_two_paragraphs_spans_them():
    package = two_paragraphs("Q", "Z")
    q, z = package.body.paragraphs
    [mark] = package.body.search("^13", **WILD)
    assert (mark.paragraph == q, mark.start, mark.end, mark.mark, mark.holds_end_mark, mark.spans_paragraphs) == (
        True, 1, 1, True, True, False,
    )
    [held] = package.body.search("Q^13", **WILD)
    assert (held.text, held.holds_end_mark, held.spans_paragraphs) == ("Q\r", True, False)
    [across] = package.body.search("Q?Z", **WILD)
    assert (across.paragraph == q, across.end_paragraph == z, across.start, across.end, across.mark) == (True, True, 0, 1, False)
    assert across.spans_paragraphs and len(across.paragraphs) == 2


def test_no_wildcard_matches_a_cells_or_a_rows_end_nor_across_a_tables_edge_a_mark_before_a_table_is_a_paragraphs():
    package = create_package()
    body = package.body
    body.insert_paragraph("B", location="End")
    table = body.insert_table(1, 2, values=[["T", "U"]], location="End")
    body.insert_paragraph("A", location="End")
    assert texts(body.search("B?T", **WILD)) == []
    assert texts(body.search("U?A", **WILD)) == []
    assert texts(body.search("T?U", **WILD)) == []
    # B's mark, before the table, is a paragraph's mark (check 59 case 2's BEGIN2); the cells' ends are no marks
    assert texts(body.search("?", **WILD)) == ["B", "\r", "T", "U", "A"]
    assert texts(table.get_range("Whole").search("?", **WILD)) == ["T", "U"]
    assert texts(table.cell(0, 0).body.search("?", **WILD)) == ["T"]


def test_a_deleted_mark_is_no_mark_to_a_wildcard_the_paragraphs_joined_as_the_default_view_reads_them():
    package = create_package()
    body = package.body
    body.insert_xml(
        f'<w:p xmlns:w="{W}"><w:pPr><w:rPr><w:del w:id="1" w:author="A" w:date="2026-10-24T00:00:00Z"/></w:rPr></w:pPr>'
        "<w:r><w:t>ONE</w:t></w:r></w:p>",
        location="End",
    )
    body.insert_paragraph("TWO", location="End")
    one, _two = body.paragraphs
    assert one.mark_text == ""
    # ^p's rule (CR-016 Phase B): a deleted mark is nothing in Word's default view, so no ^13 there
    assert texts(body.search("^13", **WILD)) == []
    assert texts(one.search("?", **WILD)) == ["O", "N", "E"]
    # the text a hit reads runs on across it, as Body.text does (ONETWO): a hit across it spans the two paragraphs and
    # holds no mark (core-ts's wildcardMatchesAcross; Word unmeasured), and a replacement of it is refused
    [hit] = body.search("E?W", **WILD)
    assert (hit.text, hit.spans_paragraphs, hit.holds_end_mark) == ("ETW", True, False)
    with pytest.raises(SpanError) as refused:
        body.replace_text("E?W", "X", **WILD)
    assert refused.value.code == "range.replaces_mark"


def test_hidden_text_is_not_read_across_a_mark_either():
    package = create_package()
    body = package.body
    body.insert_xml(
        f'<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>H</w:t></w:r>'
        "<w:r><w:t>B</w:t></w:r></w:p>",
        location="End",
    )
    body.insert_paragraph("C", location="End")
    assert texts(body.search("B?C", **WILD)) == ["B\rC"]
    assert texts(body.search("A?^13", **WILD)) == ["AB\r"]
    assert texts(body.search("H", **WILD)) == []


def test_a_block_controls_search_reads_its_content_range_its_last_mark_left_out():
    package = two_paragraphs("Q", "Z", "W")
    body = package.body
    control = body.paragraphs[0].get_range("Whole").expand_to(body.paragraphs[1].get_range("Whole")).insert_content_control()
    assert texts(control.search("?", **WILD)) == ["Q", "\r", "Z"]
    assert texts(control.search("Q?Z", **WILD)) == ["Q\rZ"]
    assert texts(control.search("Z?W", **WILD)) == []


# ---------------------------------------------------------------------------
# replace_text with wildcards (Q9)
# ---------------------------------------------------------------------------


def _refused_replacements(body, q, z):
    """Each a wildcard replacement a hit of which holds a mark: the body's, a paragraph's, a range's."""
    return (
        lambda: body.replace_text("Q?Z", "X", **WILD),
        lambda: body.replace_text("[!A-Z]", "X", **WILD),
        lambda: q.replace_text("?", "X", **WILD),
        lambda: q.replace_text("Q^13", "X", **WILD),
        lambda: body.get_range("Whole").replace_text("?", "X", **WILD),
        lambda: q.get_range("Whole").replace_text("?", "X", **WILD),
        lambda: q.get_range("Content").expand_to(z.get_range("Content")).replace_text("Q?Z", "X", **WILD),
    )


def test_a_wildcard_replacement_holding_a_mark_is_refused_before_anything_changes_in_every_scope_tracked_or_not():
    for tracked in (False, True):
        package = two_paragraphs("Q", "Z")
        if tracked:
            package.change_tracking_mode = "TrackAll"
        body = package.body
        before = body.get_xml()
        q, z = body.paragraphs
        for replace in _refused_replacements(body, q, z):
            with pytest.raises(SpanError, match="holds a paragraph mark") as refused:
                replace()
            assert refused.value.code == "range.replaces_mark"
            assert "search for the hits" in refused.value.hint
            assert body.get_xml() == before, tracked
        # a hit holding no mark is replaced, and a limit leaves the refusal reading every hit
        assert body.replace_text("Q", "X", **WILD) == 1
        with pytest.raises(SpanError):
            body.replace_text("?", "Y", limit=1, **WILD)
        assert [p.text for p in body.paragraphs] == ["X", "Z"]


def test_a_body_replacement_reads_its_content_range_so_the_last_mark_is_no_hit_where_the_paragraphs_is():
    package = two_paragraphs("Q")
    body = package.body
    paragraph = body.paragraphs[0]
    with pytest.raises(SpanError):
        paragraph.replace_text("?", "X", **WILD)
    # core-ts's Body.replaceText: the body's content holds no mark, so each paragraph's own matches are replaced
    assert body.replace_text("?", "X", **WILD) == 1
    assert paragraph.text == "X"
    # a range within one paragraph holding no mark replaces as before
    assert paragraph.get_range("Content").replace_text("X", "Y", **WILD) == 1
    assert paragraph.text == "Y"


def test_a_range_across_paragraphs_replaces_each_paragraphs_hits_when_none_holds_a_mark():
    package = two_paragraphs("Qa", "aZ")
    q, z = package.body.paragraphs
    across = q.get_range("Content").expand_to(z.get_range("Content"))
    assert texts(across.search("a", **WILD)) == ["a", "a"]
    assert across.replace_text("a", "b", **WILD) == 2
    assert [p.text for p in package.body.paragraphs] == ["Qb", "bZ"]


# ---------------------------------------------------------------------------
# the grapheme filter (Q8; core-ts ebb218e) in search, find and replace_text
# ---------------------------------------------------------------------------


def test_a_hit_inside_a_decomposed_accent_is_dropped_in_search_find_and_replace_text():
    package = two_paragraphs("cafe\u0301 cafe")
    body = package.body
    paragraph = body.paragraphs[0]
    # "cafe" ends inside the first cluster (e and U+0301): only the second is a hit
    assert [(hit.start, hit.end) for hit in paragraph.search("cafe")] == [(6, 10)]
    assert [(hit.start, hit.end) for hit in body.search("cafe")] == [(6, 10)]
    assert [(hit.start, hit.end) for hit in paragraph.get_range("Content").search("e")] == [(9, 10)]
    assert [(hit.start, hit.end) for hit in body.search("cafe", **WILD)] == [(6, 10)]
    assert [hit.match for hit in body.find("cafe")] == ["cafe"]
    assert [hit.start for hit in paragraph.find("e")] == [9]
    # the whole cluster is a hit
    assert [(hit.start, hit.end) for hit in body.search("e\u0301")] == [(3, 5)]
    assert body.replace_text("cafe", "X") == 1
    assert paragraph.text == "cafe\u0301 X"
    assert paragraph.replace_text("e", "E") == 0, "the e under the accent is not cut off it"
    assert paragraph.text == "cafe\u0301 X" and body.replace_text("cafe\u0301", "Y") == 1
    assert paragraph.text == "Y X"


def test_the_arabic_fathas_are_never_cut_off_their_letters():
    package = two_paragraphs("\u0643\u064e\u062a\u064e\u0628\u064e")
    body = package.body
    paragraph = body.paragraphs[0]
    content = paragraph.get_range("Content")
    # check 58 case 13: Word's ? and [!a-z ] find nothing in the word; nor here, through the filter
    assert texts(content.search("?", **WILD)) == []
    assert texts(content.search("[!a-z ]", **WILD)) == []
    assert texts(content.search("??", **WILD)) == ["\u0643\u064e", "\u062a\u064e", "\u0628\u064e"]
    assert texts(paragraph.search("\u062a")) == []
    assert texts(paragraph.search("\u062a\u064e")) == ["\u062a\u064e"]
    assert body.find("\u062a") == []
    assert body.replace_text("\u062a", "X") == 0
    assert paragraph.text == "\u0643\u064e\u062a\u064e\u0628\u064e"
    assert body.replace_text("\u062a\u064e", "X") == 1
    assert paragraph.text == "\u0643\u064eX\u0628\u064e"


def test_latin_text_is_untouched_the_filter_never_reads_a_boundary_below_u_0300(monkeypatch):
    def boundary(text, index):  # pragma: no cover - the test is that it is never called
        raise AssertionError(f"is_grapheme_boundary read over {text!r}")

    monkeypatch.setattr(text_model, "is_grapheme_boundary", boundary)
    package = two_paragraphs("The cat sat on the mat. caf\u00e9 \u02ff.", "Then the end.")
    body = package.body
    assert texts(body.search("the")) == ["The", "the", "The", "the"]
    assert texts(body.search("?at", **WILD)) == ["cat", "sat", "mat"]
    assert len(body.find("e")) == 5
    assert texts(body.search("[.]^13", **WILD)) == [".\r"]
    assert body.replace_text("cat", "dog") == 1
    # one code point at U+0300 or above, and the boundaries are read
    monkeypatch.undo()
    calls = []
    real = text_model.is_grapheme_boundary
    monkeypatch.setattr(text_model, "is_grapheme_boundary", lambda text, index: calls.append(index) or real(text, index))
    assert find_all("e\u0300e", search_pattern("e")) == [(2, 3)]
    assert calls


def test_find_reads_a_paragraph_at_a_time_so_a_wildcard_matches_no_mark_there():
    """``find``'s hits are one paragraph's, at an address: no mark, where ``search`` reads Word's (CR-016 section 12)."""
    package = two_paragraphs("Q", "Z")
    body = package.body
    assert [hit.match for hit in body.find("?", **WILD)] == ["Q", "Z"]
    assert body.find("Q?Z", **WILD) == []
    assert [hit.match for hit in body.paragraphs[0].find("?", **WILD)] == ["Q"]
    assert texts(body.search("?", **WILD)) == ["Q", "\r", "Z"]
