"""The writes of a Range across paragraphs: CR-013 Phase C's own tests (section 6.3, as section 13 amends it).

    .venv-fork/bin/python -m pytest tests/content/test_span_edits.py

``test_ranges.py`` ports docx4j-core-ts's ``ranges.test.mjs`` and holds the writes
to Word 365's checks 52 and 54; these are this package's own: every refusal
before anything changes, untracked and tracked, a table, a control, a kept
``mc:AlternateContent`` branch, an ``w:altChunk``; Reject All after each
tracked write giving back the original and Accept All the untracked write; the
range a write leaves (section 4.4); a held mark, a cell's and a story's last
paragraph; a whole control deleted tracked, listed and resolved; a control
wrapped around paragraphs, tracked; a section break whose relationships a join
releases; ``insert_ooxml`` and ``style`` across paragraphs; ``get_xml`` and
``to_api_script`` of a range across paragraphs (core-ts ``202a01d``'s
``office-js.test.mjs`` test); and one ``ChangeReport`` per write, every write
undone in a ``dry_run`` (CR-003 sections 3.4 and 12.5).
"""

from __future__ import annotations

import io
import re

import pytest
from conftest import ROOT

from docx4j_py import create_package, load
from docx4j_py.model import revisions
from docx4j_py.model.content import AddressError, ContentError, Range, SpanError

OFFICE_JS = ROOT / "tests" / "office_js"
FIXTURES = ROOT / "tests" / "fixtures"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def loaded(name: str = "52-ranges.docx"):
    return load(OFFICE_JS / name)


def word_in(package, token: str) -> Range:
    return package.body.search(token, match_case=True, match_whole_word=True)[0]


def across(package, start: str, end: str) -> Range:
    return word_in(package, start).expand_to(word_in(package, end))


def to_mark(package, start: str, end: str) -> Range:
    return word_in(package, start).expand_to(word_in(package, end).paragraphs[0].get_range("Whole"))


def reloaded(package):
    return load(io.BytesIO(package.save()))


def region(package, n: int) -> list:
    paragraphs = package.body.paragraphs
    begin = next(i for i, p in enumerate(paragraphs) if p.text.startswith(f"BEGIN{n} "))
    end = next(i for i, p in enumerate(paragraphs) if i > begin and p.text == f"END{n}")
    return paragraphs[begin + 1 : end]


def shape(package, n: int) -> list[tuple[str, str, str]]:
    """A region's paragraphs as the sweep compares them: text, style, alignment."""
    return [(p.text, p.style_built_in, p.alignment) for p in region(package, n)]


def document_xml(package) -> str:
    return package.main_document_part.xml.decode()


#: The six writes of the sweep (CR-013 section 6.3), by the check 52 region whose range they take (a heading's
#: second word to a body paragraph's first: the HB kind), and a 'Whole' range's deletion (W1).
WRITES = {
    "delete": (3, lambda r: r.delete()),
    "replace": (29, lambda r: r.insert_text("NEW", location="Replace")),
    "bold": (49, lambda r: setattr(r.font, "bold", True)),
    "paragraph-before": (57, lambda r: r.insert_paragraph("NEW", location="Before")),
    "paragraph-after": (63, lambda r: r.insert_paragraph("NEW", location="After")),
}


def hb_range(package, n: int) -> Range:
    return across(package, f"HEAD{n}B", f"BODY{n}A")


@pytest.mark.parametrize("write", list(WRITES), ids=list(WRITES))
def test_accept_all_of_a_tracked_write_gives_the_untracked_write_and_reject_all_gives_back_the_original(write):
    n, apply = WRITES[write]
    untracked = loaded()
    apply(hb_range(untracked, n))
    untracked = reloaded(untracked)
    revisions.accept_all(untracked)
    tracked = loaded()
    tracked.change_tracking_mode = "TrackAll"
    apply(hb_range(tracked, n))
    tracked.change_tracking_mode = "Off"
    saved = tracked.save()
    accepted, rejected = load(io.BytesIO(saved)), load(io.BytesIO(saved))
    revisions.accept_all(accepted)
    revisions.reject_all(rejected)
    original = loaded()
    revisions.reject_all(original)
    assert shape(accepted, n) == shape(untracked, n)
    assert shape(rejected, n) == shape(original, n)


def test_a_whole_paragraphs_range_tracked_and_untracked_accept_gives_untracked_reject_gives_back():
    """W1 (check 52 case 14): a paragraph's 'Whole' range deletes the paragraph, joining the next; tracked its mark."""
    untracked = loaded()
    paragraphs = len(untracked.body.paragraphs)
    word_in(untracked, "BODY14A").paragraphs[0].get_range("Whole").delete()
    assert len(untracked.body.paragraphs) == paragraphs - 1
    assert [p.text for p in region(untracked, 14)] == [t for t in [p.text for p in region(loaded(), 14)] if "BODY14A" not in t]
    tracked = loaded()
    tracked.change_tracking_mode = "TrackAll"
    word_in(tracked, "BODY14A").paragraphs[0].get_range("Whole").delete()
    tracked.change_tracking_mode = "Off"
    assert len(tracked.body.paragraphs) == paragraphs, "tracked, the paragraph stays until accepted"
    saved = tracked.save()
    accepted, rejected = load(io.BytesIO(saved)), load(io.BytesIO(saved))
    revisions.accept_all(accepted)
    revisions.reject_all(rejected)
    accepted_untracked = reloaded(untracked)
    revisions.accept_all(accepted_untracked)
    original = loaded()
    revisions.reject_all(original)
    assert shape(accepted, 14) == shape(accepted_untracked, 14)
    assert shape(rejected, 14) == shape(original, 14)


def test_the_range_a_write_leaves():
    package = loaded()
    # a deletion: collapsed where the paragraphs joined
    deleted = hb_range(package, 3)
    deleted.delete()
    assert [deleted.spans_paragraphs, deleted.text, deleted.start, deleted.paragraph.text] == [False, "", 7, "HEAD3A  BODY3B"]
    # a replacement: over the new text, its mark no longer held
    replaced = to_mark(package, "HEAD38B", "BODY38A")
    replaced.insert_text("NEW", location="Replace")
    assert [replaced.text, replaced.mark, replaced.holds_end_mark] == ["NEW", False, False]
    # a paragraph inserted at the start splits there: the range keeps its text, from the start of what was split off
    split = hb_range(package, 57)
    split.insert_paragraph("NEW", location="Before")
    assert split.text == "HEAD57B\rBODY57A"
    # and at the end: the text before the split moved into the first half, which the range now ends in
    after = hb_range(package, 63)
    after.insert_paragraph("NEW", location="After")
    assert after.text == "HEAD63B\rBODY63A"
    assert after.end_paragraph.text == "BODY63A"
    # text at the ends: the range keeps covering its own text
    ends = hb_range(package, 71)
    ends.insert_text("<", location="Start")
    ends.insert_text(">", location="End")
    assert ends.text == "HEAD71B\rBODY71A"
    assert [p.text for p in ends.paragraphs] == ["HEAD71A <HEAD71B", "BODY71A> BODY71B"]
    # tracked, a deletion leaves the range collapsed where it starts, as Simple Markup reads it
    package.change_tracking_mode = "TrackAll"
    tracked = hb_range(package, 16)
    tracked.delete()
    assert [tracked.spans_paragraphs, tracked.text, tracked.start] == [False, "", 8]


def test_a_held_mark_with_no_paragraph_after_it():
    """check 54 cases 8, 10 and 12, and a story's last paragraph: the mark goes, or is kept."""
    package = loaded("54-writes.docx")
    body = package.body
    # before a table (case 8): the paragraph goes
    count = len(body.paragraphs)
    word_in(package, "BODY8A").paragraphs[0].get_range("Whole").delete()
    assert len(body.paragraphs) == count - 1
    assert [p.text for p in region(package, 8)][0] == "T8R1C1A T8R1C1B"
    # a cell's last paragraph (case 10): emptied and kept
    word_in(package, "LIST10A").paragraphs[0].get_range("Whole").delete()
    assert [p.text for p in region(package, 10)] == ["BODY10A BODY10B", "", "T10R1C2A T10R1C2B"]
    # a control's last paragraph (case 12): the paragraph goes
    control = next(c for c in body.content_controls if c.tag == "c54-12")
    word_in(package, "LIST12A").paragraphs[0].get_range("Whole").delete()
    assert [p.text for p in control.paragraphs] == ["HEAD12A HEAD12B HEAD12C", "BODY12A BODY12B"]
    # a story's last paragraph keeps its mark (unmeasured: Word can't delete the last mark)
    created = create_package()
    created.body.insert_paragraph("only", location="End")
    last = created.body.paragraphs[-1]
    last.get_range("Whole").delete()
    assert created.body.paragraphs[-1].element is last.element and last.text == ""


def test_every_refusal_comes_before_anything_changes_untracked_and_tracked():
    package = loaded()
    body = package.body
    before = document_xml(package)
    for mode in ("Off", "TrackAll"):
        package.change_tracking_mode = mode
        cases = [
            # into and out of a table: Word's own refusal (check 52's INTO and CELL)
            (lambda: across(package, "HEAD9B", "T9R2C1A").delete(), SpanError, "range.partial_table"),
            (lambda: across(package, "T10R1C1B", "BODY10A").insert_text("X", location="Replace"), SpanError, "range.partial_table"),
            # a replacement ending with a whole table (check 54 case 7)
            (
                lambda: word_in(package, "HEAD21B").expand_to(word_in(package, "T21R1C1A").parent_table.get_range()).insert_text("X", location="Replace"),
                SpanError,
                "range.replaces_table",
            ),
            # a replacement across a control's edge, unmeasured
            (
                lambda: word_in(package, "BEGIN13").expand_to(next(c for c in body.content_controls if c.tag == "c52-13").paragraphs[0].search("HEAD13A")[0]).insert_text("X", location="Replace"),
                SpanError,
                "range.partial_control",
            ),
            # a control around a range not holding its paragraphs whole (check 54 cases 36, 37)
            (lambda: hb_range(package, 3).insert_content_control(), SpanError, "range.partial_paragraph"),
            # a control around a range across a table (check 54 case 39)
            (lambda: to_mark(package, "HEAD21A", "BODY21A").expand_to(word_in(package, "HEAD21A").paragraphs[0].get_range("Start")).insert_content_control(), SpanError, "range.partial_paragraph"),
            # ^p in a replacement
            (lambda: body.replace_text("^p", " "), SpanError, "range.replaces_mark"),
            # a hyperlink across paragraphs
            (lambda: setattr(hb_range(package, 3), "hyperlink", "https://example.com"), SpanError, "range.spans_paragraphs"),
            # several blocks in place of a range across paragraphs, unmeasured
            (
                lambda: hb_range(package, 3).insert_ooxml(f'<w:p xmlns:w="{W}"/><w:p xmlns:w="{W}"/>', location="Replace"),
                SpanError,
                "range.spans_paragraphs",
            ),
        ]
        for write, kind, code in cases:
            with pytest.raises(kind) as refused:
                write()
            assert refused.value.code == code, (mode, code)
    package.change_tracking_mode = "Off"
    assert document_xml(package) == before


def test_a_whole_alt_chunk_is_not_trackable_and_goes_untracked():
    package = create_package()
    body = package.body
    body.insert_paragraph("HEAD one", location="End")
    body.insert_xml(f'<w:altChunk xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdChunk"/>', location="End")
    body.insert_paragraph("BODY two", location="End")
    package.change_tracking_mode = "TrackAll"
    span = body.search("one")[0].expand_to(body.search("BODY")[0])
    before = document_xml(package)
    with pytest.raises(SpanError) as refused:
        span.delete()
    assert refused.value.code == "range.not_trackable" and "altChunk" in str(refused.value)
    assert document_xml(package) == before
    package.change_tracking_mode = "Off"
    span.delete()
    assert "altChunk" not in document_xml(package)
    assert [p.text for p in body.paragraphs if p.text] == ["HEAD  two"]


def test_a_kept_branch_is_read_only_a_range_into_one_refuses_and_one_holding_its_text_refuses():
    """CR-013 section 4.5: no editing in place through a kept branch, a departure from core-ts's ED-009."""
    block = load(FIXTURES / "mc-fun" / "mc-block.docx")
    first, kept = block.body.paragraphs[0], block.body.paragraphs[-1]
    assert kept.text == "w14"
    span = first.get_range("Start").expand_to(kept.get_range("End"))
    before = block.main_document_part.xml
    for write in (span.delete, lambda: span.insert_text("X", location="End"), lambda: setattr(span.font, "bold", True)):
        with pytest.raises(AddressError) as refused:
            write()
        assert refused.value.code == "address.read_only"
    assert block.main_document_part.xml == before
    run = load(FIXTURES / "mc-fun" / "mc-run.docx")
    paragraphs = run.body.paragraphs
    branch = next(p for p in paragraphs if "w14" in p.text)
    # the range's head piece holds the branch's text: refused, untracked and tracked
    span = branch.search("prefer")[0].expand_to(paragraphs[-1].get_range("End"))
    for mode in ("Off", "TrackAll"):
        run.change_tracking_mode = mode
        with pytest.raises(AddressError) as refused:
            span.delete()
        assert refused.value.code == "address.read_only"
    run.change_tracking_mode = "Off"
    # a paragraph holding the branch's text between the ends goes whole untracked, as Paragraph.delete() takes it
    whole = paragraphs[0].get_range("Start").expand_to(paragraphs[-1].get_range("End"))
    whole.delete()
    assert "w14" not in "".join(p.text for p in run.body.paragraphs)


def test_a_whole_control_deleted_tracked_is_listed_deleted_and_resolves_as_word_resolves_it():
    package = loaded("54-writes.docx")
    package.change_tracking_mode = "TrackAll"
    hb = across(package, "HEAD15B", "BODY15A")
    hb.delete()
    package.change_tracking_mode = "Off"
    saved = package.save()
    listed = [c for c in revisions.tracked_changes(load(io.BytesIO(saved))) if "PLAIN15A" in c.text or "HEAD15B" in c.text]
    assert len(listed) == 1 and listed[0].type == "Deleted"
    pieces = [piece.kind for piece in listed[0].pieces]
    assert "content_control_deletion" in pieces
    accepted, rejected = load(io.BytesIO(saved)), load(io.BytesIO(saved))
    revisions.accept_all(accepted)
    revisions.reject_all(rejected)
    assert [c for c in accepted.body.content_controls if c.tag == "c54-15"] == []
    assert [p.text for p in region(accepted, 15)] == ["HEAD15A  BODY15B"]
    assert [c.tag for c in rejected.body.content_controls if c.tag == "c54-15"] == ["c54-15"]
    assert "customXmlDelRange" not in rejected.main_document_part.xml.decode()
    original = loaded("54-writes.docx")
    assert shape(rejected, 15) == shape(original, 15)


def test_a_control_wrapped_around_paragraphs_tracked_is_an_insertion_reject_all_unwraps_it():
    package = loaded("54-writes.docx")
    package.change_tracking_mode = "TrackAll"
    span = word_in(package, "HEAD43A").paragraphs[0].get_range("Start").expand_to(word_in(package, "BODY43A").paragraphs[0].get_range("Whole"))
    control = span.insert_content_control()
    package.change_tracking_mode = "Off"
    assert control.form == "block" and [p.text for p in control.paragraphs] == ["HEAD43A HEAD43B HEAD43C", "BODY43A BODY43B"]
    assert re.search(r"<w:customXmlInsRangeStart[^>]*/><w:sdt>", package.main_document_part.xml.decode())
    saved = package.save()
    accepted, rejected = load(io.BytesIO(saved)), load(io.BytesIO(saved))
    revisions.accept_all(accepted)
    revisions.reject_all(rejected)
    assert "customXmlInsRange" not in accepted.main_document_part.xml.decode()
    assert len(accepted.body.content_controls) == len(package.body.content_controls)
    assert len(rejected.body.content_controls) == len(loaded("54-writes.docx").body.content_controls)
    assert [p.text for p in region(rejected, 43)] == ["HEAD43A HEAD43B HEAD43C", "BODY43A BODY43B", "PLAIN43A PLAIN43B"]


def test_a_section_break_in_the_first_mark_goes_with_it_its_headers_released():
    """core-ts CR-014 section 10.1: a section break in the first paragraph's mark goes with the mark, and the
    relationships its headers and footers held are released (here when the write ends, through Removals)."""
    package = load(FIXTURES / "header-no-rels.docx")
    paragraphs = package.body.paragraphs
    broken = next(p for p in paragraphs if p.element.p_pr is not None and p.element.p_pr.sect_pr is not None)
    rels = package.main_document_part.relationships_part
    assert {rel.id for rel in rels.list if rel.type_value.endswith("/header")} == {"rId6", "rId7", "rId8"}
    Range(broken, 0, 0).expand_to(package.body.search("Some")[0]).delete()
    assert [p.text for p in package.body.paragraphs][-1].startswith(" content in the second section")
    assert not any(p.element.p_pr is not None and p.element.p_pr.sect_pr is not None for p in package.body.paragraphs)
    # its two headers' relationships and parts went; the body's own section keeps its header
    assert {rel.id for rel in rels.list if rel.type_value.endswith("/header")} == {"rId8"}
    assert sorted(str(part.part_name) for part in package.parts.parts() if "header" in str(part.part_name)) == ["/word/header3.xml"]
    back = reloaded(package)
    assert sorted(str(part.part_name) for part in back.parts.parts() if "header" in str(part.part_name)) == ["/word/header3.xml"]


def test_style_across_paragraphs_sets_each_tracked_with_a_ppr_change_each():
    package = loaded("54-writes.docx")
    package.change_tracking_mode = "TrackAll"
    across(package, "BODY33B", "LIST33A").style_built_in = "Heading2"
    package.change_tracking_mode = "Off"
    styled = region(package, 33)
    assert [p.style_built_in for p in styled] == ["Heading2", "Heading2"]
    xml = package.main_document_part.xml.decode()
    part = xml[xml.index("BEGIN33 ") : xml.index("END33")]
    assert part.count("<w:pPrChange") == 2
    rejected = reloaded(package)
    revisions.reject_all(rejected)
    assert [p.style_built_in for p in region(rejected, 33)] == ["Normal", "ListParagraph"]
    untracked = loaded("54-writes.docx")
    across(untracked, "BODY32B", "LIST32A").style = "Heading 2"
    assert [p.style_built_in for p in region(untracked, 32)] == ["Heading2", "Heading2"]


def test_insert_ooxml_across_paragraphs():
    """check 54 case 35: one paragraph's runs in place of the range; Before and After at its ends; blocks beside it."""
    one = f'<w:p xmlns:w="{W}"><w:r><w:t>NEW</w:t></w:r></w:p>'
    two = f'<w:p xmlns:w="{W}"><w:r><w:t>ONE</w:t></w:r></w:p><w:p xmlns:w="{W}"><w:r><w:t>TWO</w:t></w:r></w:p>'
    package = loaded()
    (into,) = hb_range(package, 3).insert_ooxml(one, location="Replace")
    assert into.text == "HEAD3A NEW BODY3B"
    hb_range(package, 16).insert_ooxml(one, location="Before")
    hb_range(package, 39).insert_ooxml(one, location="After")
    assert [p.text for p in region(package, 16)] == ["HEAD16A NEWHEAD16B", "BODY16A BODY16B"]
    assert [p.text for p in region(package, 39)] == ["HEAD39A HEAD39B", "BODY39ANEW BODY39B"]
    views = hb_range(package, 21).insert_ooxml(two, location="After")
    assert [v.text for v in views] == ["ONE", "TWO"]
    assert [p.text for p in region(package, 21)][-3:] == ["BODY21A BODY21B", "ONE", "TWO"]


def test_get_xml_and_to_api_script_of_a_range_across_paragraphs_give_its_blocks_cut_untracked_from_copies():
    """core-ts 202a01d's office-js.test.mjs test, for get_xml and to_api_script."""
    package = create_package()
    body = package.body
    body.insert_paragraph("alpha beta gamma", location="End")
    body.insert_table(1, 2, location="End", values=[["c1", "c2"]])
    body.insert_paragraph("delta epsilon", location="End")
    before = document_xml(package)

    def hit(word: str) -> Range:
        return body.search(word, match_case=True)[0]

    def texts(xml: str) -> list[str]:
        return re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", xml)

    # within one paragraph, the paragraph trimmed to the range
    assert texts(hit("beta").get_xml()) == ["beta"]
    # across paragraphs: the first's tail, the table whole, the last's head
    xml = hit("beta").expand_to(hit("delta")).get_xml()
    assert re.match(r"^<w:p [\s\S]*?</w:p><w:tbl[\s\S]*</w:tbl><w:p [\s\S]*</w:p>$", xml)
    assert texts(xml) == ["beta gamma", "c1", "c2", "delta"]
    # which insert_xml takes back
    target = create_package()
    target.body.insert_xml(xml, location="End")
    assert target.body.text.endswith("beta gamma\rc1\tc2\r\ndelta")
    # a range starting or ending inside a table it doesn't hold whole is refused
    with pytest.raises(SpanError) as refused:
        hit("beta").expand_to(hit("c1")).get_xml()
    assert refused.value.code == "range.partial_table" and "inside a table" in str(refused.value)
    # under tracking the copies are cut, not marked deleted, and the document is untouched (a tracked get_xml of one
    # word gave the paragraph's other words as w:del until CR-013 Phase C, core-ts's defect, which was here too)
    package.change_tracking_mode = "TrackAll"
    assert "w:del" not in hit("beta").get_xml()
    tracked = hit("gamma").expand_to(hit("delta")).get_xml()
    assert "w:del" not in tracked and texts(tracked) == ["gamma", "c1", "c2", "delta"]
    package.change_tracking_mode = "Off"
    assert re.sub(r"<w:trackRevisions[^>]*/>", "", document_xml(package)) == before
    # to_api_script over the same range writes each block, against a body
    script = hit("beta").expand_to(body.paragraphs[-1].search("delta")[0]).to_api_script()
    assert script.startswith('body.insert_paragraph("beta gamma")')
    assert "delta" in script
    replayed = create_package()
    exec(script, {"body": replayed.body})  # noqa: S102 - the script is the API's own output
    assert replayed.body.text.endswith("beta gamma\rc1\tc2\r\ndelta")


def test_a_comment_whose_reference_a_deletion_across_paragraphs_takes_goes():
    package = loaded("54-writes.docx")
    names = {c.content for c in package.body.get_comments()}
    assert "INSIDE27" in names
    across(package, "HEAD27A", "BODY27A").delete()
    assert "INSIDE27" not in {c.content for c in package.body.get_comments()}
    assert "INSIDE27" not in {c.content for c in reloaded(package).body.get_comments()}
    # one whose reference stays keeps its markers at the join (case 29), as a bookmark does (case 30)
    across(package, "HEAD29A", "BODY29A").delete()
    kept = next(c for c in package.body.get_comments() if c.content == "ACROSS29")
    assert kept.get_range().text == " BODY29B"


def test_the_facades_add_comment_still_refuses_runs_in_two_paragraphs():
    from docx4j_py.docx import Document

    document = Document()
    first = document.add_paragraph("one")
    second = document.add_paragraph("two")
    with pytest.raises(ContentError) as refused:
        document.add_comment([first.runs[0], second.runs[0]], text="across")
    assert refused.value.code == "comment.spans_paragraphs"


def _one_report(package, write) -> object:
    count = len(package.changes)
    write()
    assert len(package.changes) == count + 1, "one ChangeReport a write"
    return package.changes[-1]


DRY_WRITES = {
    "delete": lambda p: hb_range(p, 3).delete(),
    "delete a note": lambda p: across(p, "HEAD12B", "BODY12A").delete(),
    "replace": lambda p: hb_range(p, 29).insert_text("NEW", location="Replace"),
    "bold": lambda p: setattr(hb_range(p, 49).font, "bold", True),
    "paragraph-before": lambda p: hb_range(p, 57).insert_paragraph("NEW", location="Before"),
    "text-end": lambda p: hb_range(p, 75).insert_text("NEW", location="End"),
    "replace_text": lambda p: hb_range(p, 3).replace_text("3", "X"),
    "style": lambda p: setattr(hb_range(p, 3), "style_built_in", "Heading2"),
    "comment": lambda p: hb_range(p, 80).insert_comment("C"),
    "ooxml": lambda p: hb_range(p, 3).insert_ooxml(f'<w:p xmlns:w="{W}"><w:r><w:t>N</w:t></w:r></w:p>', location="Replace"),
    "control": lambda p: word_in(p, "HEAD3A").paragraphs[0].get_range("Start").expand_to(word_in(p, "BODY3A").paragraphs[0].get_range("Whole")).insert_content_control(),
    "whole": lambda p: word_in(p, "BODY14A").paragraphs[0].get_range("Whole").delete(),
}

OPERATIONS = {
    "delete": "delete",
    "delete a note": "delete",
    "replace": "insert_text",
    "bold": "format",
    "paragraph-before": "insert_paragraph",
    "text-end": "insert_text",
    "replace_text": "replace_text",
    "style": "format",
    "comment": "insert_comment",
    "ooxml": "insert_ooxml",
    "control": "insert_content_control",
    "whole": "delete",
}


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("write", list(DRY_WRITES), ids=list(DRY_WRITES))
def test_each_write_records_one_change_report_and_a_dry_run_leaves_the_package_as_it_was(write, tracked):
    package = loaded()
    package.get_run_font_selector()
    if tracked:
        package.change_tracking_mode = "TrackAll"
    before = {str(part.part_name): part.xml for part in package.parts.parts() if getattr(part, "is_unmarshalled", False)}
    with package.dry_run() as trial:
        DRY_WRITES[write](trial)
        assert len(trial.changes) == 1
        report = trial.changes[-1]
        assert report.operation == OPERATIONS[write]
    after = {str(part.part_name): part.xml for part in package.parts.parts() if getattr(part, "is_unmarshalled", False)}
    assert after == before, "the real package is what it was"
    real = _one_report(package, lambda: DRY_WRITES[write](package))
    assert real.operation == OPERATIONS[write]
    assert real.addresses, "the paragraphs touched"
    assert "/word/document.xml" in real.parts_touched
    if write == "delete a note" and not tracked:
        # a note whose reference went leaves the notes part, in the dry run's report too (tracked it stays, in the w:del)
        assert "/word/footnotes.xml" in real.parts_touched and "/word/footnotes.xml" in report.parts_touched
    if write in ("delete", "replace"):
        assert real.text_before == "HEAD3B\rBODY3A" if write == "delete" else real.text_before == "HEAD29B\rBODY29A"
    reloaded(package)


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_a_range_of_one_table_taken_whole_takes_that_table_only(tracked):
    # found by the sweep (CR-013 section 14.9): a range holding one block whole removed the block after it too
    package = create_package()
    package.body.insert_paragraph("before")
    package.body.insert_table(2, 2, values=[["a", "b"], ["c", "d"]])
    package.body.insert_paragraph("after")
    package.body.insert_paragraph("last")
    if tracked:
        package.change_tracking_mode = "TrackAll"
    first, last = (next(p for p in package.body.paragraphs if p.text == text) for text in ("a", "d"))
    span = first.get_range("Start").expand_to(last.get_range("Whole"))
    assert span.text == "a\tb\r\nc\td\t"
    span.delete()
    if tracked:
        assert document_xml(package).count("<w:delText>") == 4
        revisions.accept_all(package)
    assert [p.text for p in package.body.paragraphs] == ["before", "after", "last"]
    assert [type(b).__name__ for b in package.main_document_part.contents.body.content] == ["P", "P", "P"]


def test_a_tracked_replacement_inside_another_authors_insertion_splits_it_the_new_text_between():
    # CR-013 section 13.1 item 7: CR-003's rule, the deletion now in the half after the new text
    package = create_package()
    package.body.insert_xml(
        f'<w:p xmlns:w="{W}"><w:r><w:t xml:space="preserve">a </w:t></w:r>'
        '<w:ins w:id="7" w:author="Other" w:date="2026-01-01T00:00:00Z"><w:r><w:t>INSERTED</w:t></w:r></w:ins>'
        '<w:r><w:t xml:space="preserve"> b</w:t></w:r></w:p>',
        location="End",
    )
    package.change_tracking_mode = "TrackAll"
    paragraph = next(p for p in package.body.paragraphs if p.text)
    out = Range(paragraph, 6, 7).insert_text("X", location="Replace")
    assert out.text == "X"
    xml = to_xml_of(paragraph)
    authors = re.findall(r'<w:(ins|del) w:id="\d+" w:author="([^"]+)"', xml)
    assert authors == [("ins", "Other"), ("ins", "docx4j-python"), ("ins", "Other"), ("del", "docx4j-python")]
    assert re.search(r"<w:t>INSE</w:t></w:r></w:ins><w:ins [^>]*><w:r><w:t>X</w:t>", xml)
    assert re.search(r"<w:ins [^>]*Other[^>]*><w:del [^>]*><w:r><w:delText>R</w:delText></w:r></w:del><w:r><w:t>TED", xml)
    accepted, rejected = reloaded(package), reloaded(package)
    revisions.accept_all(accepted)
    revisions.reject_all(rejected)
    assert next(p for p in accepted.body.paragraphs if p.text).text == "a INSEXTED b"
    assert next(p for p in rejected.body.paragraphs if p.text).text == "a  b"


def test_a_whole_paragraphs_content_range_takes_an_empty_paragraphs_comment_reference():
    # core-ts's fix of an empty span's edges (its CR-014 section 12.6), found by the sweep over invoice.docx
    for write in (lambda r: r.delete(), lambda r: r.insert_text("", location="Replace")):
        package = create_package()
        paragraph = package.body.insert_paragraph("x")
        paragraph.get_range("Whole").insert_comment("on nothing")
        # the text gone, the comment's markup left: a paragraph holding only a reference, as invoice.docx has one
        items = paragraph.element.content
        items[:] = [item for item in items if "w:t>" not in to_xml_of_element(item)]
        assert paragraph.text == "" and "commentReference" in to_xml_of(paragraph)
        write(paragraph.get_range("Content"))
        assert "commentReference" not in to_xml_of(paragraph)
        assert package.body.get_comments() == []


def to_xml_of(paragraph) -> str:
    return to_xml_of_element(paragraph.element)


def to_xml_of_element(element) -> str:
    from docx4j_py.wml import to_xml

    return to_xml(element)
