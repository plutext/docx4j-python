"""The engine's contract and the rules no Word-made fixture holds. CR-010 section 4.

The port of docx4j-core-tests' ``TrackedChangesTest`` (docx4j CR-043 section 12)
as far as Phase A reaches: the grouping, dates, a stale view, a piece resolved
on its own, the unmeasured rules of CR-043 section 11 (U1, U2, U3, U15, a
hyperlink's relationship in a deleted row, the text box twin, an equation
structure) on handwritten markup; the table, section and cell records listed
but left for Phase B. Then what this port adds: the typed islands (CR-010
section 3.5), every story, and parts nobody reads left byte for byte.
"""

from __future__ import annotations

import datetime
import io
import zipfile

import pytest
from revision_helpers import (
    NS,
    body_xml,
    dele,
    fixture,
    ins,
    paragraphs,
    revisions,
    round_trip,
    run,
    text_of,
    unresolved,
    with_body,
)

from docx4j_py.model.content.errors import TrackedChangeError
from docx4j_py.model.revisions.lister import Lister

# ---------------------------------------------------------------------------
# listing
# ---------------------------------------------------------------------------


def test_deletions_that_touch_are_one_change():
    """Check 22: a deleted paragraph is one change, a lone deleted mark one, a deletion across a mark one."""
    package = fixture("core-ts/check22/22-deleted.docx")
    changes = revisions.tracked_changes(package)
    assert all(change.type == "Deleted" for change in changes)
    assert [change.text for change in changes] == ["Two.\r", "\r", "six.\rSeven "]


def test_another_author_or_type_splits():
    """Check 23: another author splits a group; a deletion and an insertion are two changes."""
    package = with_body(
        "<w:p>" + dele(1, "A", "one ") + dele(2, "B", "two ") + ins(3, "B", "three ") + ins(4, "B", "four") + "</w:p>"
    )
    changes = revisions.tracked_changes(package)
    assert [change.text for change in changes] == ["one ", "two ", "three four"]
    assert len(changes[2].pieces) == 2
    assert changes[2].kind == "run_insertion"


def test_the_equations_list_as_word_counts_them():
    """CR-010 section 1: the equations fixture listed 0 changes; Word counts 12."""
    assert len(revisions.tracked_changes(fixture("tracked-changes-equations.docx"))) == 12


def test_a_text_box_is_listed_once_and_resolves_its_twin():
    """C22: listed from mc:Choice; accepting it resolves its mc:Fallback twin, paired by position."""
    package = fixture("cr043/cr043-tracked.docx")
    listed = [change for change in revisions.tracked_changes(package) if "C22 inserted" in change.text]
    assert len(listed) == 1
    inserted = listed[0]
    assert len(inserted.twins) == 1
    inserted.accept()
    xml = body_xml(package)
    first = xml.index("C22 inserted")
    assert xml.index("C22 inserted", first + 1) > 0, "both branches keep the text"
    for change in revisions.tracked_changes(package):
        assert "C22 inserted" not in change.text
    for change in Lister(package, False).list():
        assert "C22 inserted" not in change.text, "and the Fallback's insertion went with it"


# ---------------------------------------------------------------------------
# dates
# ---------------------------------------------------------------------------


def test_date_utc_is_preferred():
    package = with_body(
        '<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T10:00:00Z" '
        'w16du:dateUtc="2026-01-01T00:00:00Z">' + run("x") + "</w:ins></w:p>"
    )
    assert revisions.tracked_changes(package)[0].date == datetime.datetime(2026, 1, 1, tzinfo=datetime.UTC)


def test_w_date_is_read_as_written():
    """A departure from docx4j (CR-010 section 10.3): ``w:date`` is read as UTC, as this package writes it."""
    package = with_body('<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T10:00:00Z">' + run("x") + "</w:ins></w:p>")
    assert revisions.tracked_changes(package)[0].date == datetime.datetime(2026, 1, 1, 10, tzinfo=datetime.UTC)


def test_no_date_is_none():
    package = with_body('<w:p><w:ins w:id="1" w:author="A">' + run("x") + "</w:ins></w:p>")
    assert revisions.tracked_changes(package)[0].date is None


# ---------------------------------------------------------------------------
# the contract
# ---------------------------------------------------------------------------


def test_a_view_is_stale_after_the_next_accept_or_reject():
    package = with_body("<w:p>" + ins(1, "A", "one") + run(" plain ") + ins(2, "A", "two") + "</w:p>")
    changes = revisions.tracked_changes(package)
    assert len(changes) == 2
    changes[1].accept()
    with pytest.raises(TrackedChangeError) as stale:
        changes[0].accept()
    assert stale.value.code == "tracking.gone"
    again = revisions.tracked_changes(package)
    assert len(again) == 1
    again[0].reject()
    assert text_of(paragraphs(package)[0]) == " plain two"


def test_a_content_api_view_is_stale_too():
    """CR-010 section 3.2: the content API's view raises ``tracking.gone`` as Phase F's did."""
    package = with_body("<w:p>" + ins(1, "A", "one") + run(" plain ") + ins(2, "A", "two") + "</w:p>")
    first, second = package.get_tracked_changes()
    second.accept()
    with pytest.raises(TrackedChangeError) as stale:
        first.reject()
    assert stale.value.code == "tracking.gone"
    assert "get_tracked_changes" in stale.value.hint


def test_a_piece_resolves_on_its_own():
    """A group's piece is a view of its own: rejected alone, it leaves the rest of the group pending."""
    package = with_body("<w:p>" + run("a ") + ins(1, "A", "one ") + ins(2, "A", "two") + "</w:p>")
    changes = revisions.tracked_changes(package)
    assert len(changes) == 1
    assert len(changes[0].pieces) == 2
    changes[0].pieces[0].reject()
    left = revisions.tracked_changes(package)
    assert [change.text for change in left] == ["two"]
    assert text_of(paragraphs(package)[0]) == "a two"


def test_a_filter_resolves_the_pieces_it_matches():
    """A filter applies to single changes (U19): a group partly matched has only those pieces resolved."""
    package = with_body(
        '<w:p><w:pPr><w:rPr><w:ins w:id="2" w:author="A"/></w:rPr></w:pPr>'
        + run("kept ")
        + ins(1, "A", "inserted")
        + "</w:p><w:p>"
        + run("next")
        + "</w:p>"
    )
    assert len(revisions.tracked_changes(package)) == 1
    revisions.reject_all(package, lambda piece: piece.kind == "run_insertion")
    left = revisions.tracked_changes(package)
    assert [str(change.kind) for change in left] == ["paragraph_mark_insertion"]
    assert len(paragraphs(package)) == 2

    by_author = with_body("<w:p>" + ins(1, "A", "a1 ") + run("x ") + dele(2, "A", "a2 ") + run("y ") + ins(3, "B", "b1") + "</w:p>")
    assert revisions.accept_all(by_author, lambda piece: piece.author == "A") == 2
    assert [change.author for change in revisions.tracked_changes(by_author)] == ["B"]
    assert text_of(paragraphs(by_author)[0]) == "a1 x y b1"


# ---------------------------------------------------------------------------
# what Phase A lists and leaves to Phase B
# ---------------------------------------------------------------------------

TABLE_RECORD = (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
    '<w:tblPrChange w:id="1" w:author="A"><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr></w:tblPrChange></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p>' + run("cell") + "</w:p></w:tc></w:tr></w:tbl>"
)


def test_a_tables_records_are_listed_and_left_for_phase_b():
    """docx4j phase 1: listed as one change per table, never resolved by Phase A, so never reported clean."""
    package = with_body(TABLE_RECORD + "<w:p>" + ins(2, "A", "text") + "</w:p>")
    changes = revisions.tracked_changes(package)
    assert [str(change.kind) for change in changes] == ["table_properties", "run_insertion"]
    assert changes[0].text == "cell\r\n"
    assert not changes[0].resolvable
    with pytest.raises(TrackedChangeError) as unsupported:
        changes[0].reject()
    assert unsupported.value.code == "tracking.unsupported"
    assert revisions.accept_all(package) == 1, "the count is of what was resolved"
    left = revisions.tracked_changes(package)
    assert [str(change.kind) for change in left] == ["table_properties"]
    assert "tblPrChange" in body_xml(package)


def test_a_sections_record_is_listed():
    package = with_body(
        "<w:p>" + run("text") + "</w:p><w:sectPr>"
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>'
        '<w:sectPrChange w:id="1" w:author="A"><w:sectPr>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="2160" w:header="708" w:footer="708" w:gutter="0"/>'
        "</w:sectPr></w:sectPrChange></w:sectPr>"
    )
    changes = revisions.tracked_changes(package)
    assert [str(change.kind) for change in changes] == ["section_properties"]
    assert changes[0].text == "text"
    assert package.outline().stats.tracked_changes == 1


def test_a_cell_insertion_is_listed():
    package = with_body(
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p>' + run("one") + "</w:p></w:tc>"
        '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:cellIns w:id="1" w:author="A"/></w:tcPr><w:p>'
        + run("two")
        + "</w:p></w:tc></w:tr></w:tbl><w:p/>"
    )
    assert [str(change.kind) for change in revisions.tracked_changes(package)] == ["cell_insertion"]


# ---------------------------------------------------------------------------
# the unmeasured rules (docx4j CR-043 section 11)
# ---------------------------------------------------------------------------

DELETED_MARK = '<w:pPr><w:rPr><w:del w:id="{0}" w:author="A"/></w:rPr></w:pPr>'


def test_u1_no_join_across_a_container_boundary():
    """A deleted mark with no paragraph after it in its list is dropped, with no join; nor across a table."""
    package = with_body(
        "<w:p>" + DELETED_MARK.format(1) + run("before a table") + "</w:p>"
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p>'
        + DELETED_MARK.format(2)
        + run("a cell's last")
        + "</w:p></w:tc></w:tr></w:tbl>"
        "<w:p>" + DELETED_MARK.format(3) + run("the story's last") + "</w:p><w:sectPr/>"
    )
    assert revisions.accept_all(package) == 3
    assert len(paragraphs(package)) == 2
    assert revisions.tracked_changes(package) == []
    assert "w:del " not in body_xml(package)
    assert package.main_document_part.contents.body.sect_pr is not None


def test_a_join_keeps_the_next_paragraph():
    """A deleted mark joins: the result keeps the next paragraph's properties and paraId (C10b); markers between go into it."""
    package = with_body(
        '<w:p w14:paraId="00000001"><w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:del w:id="1" w:author="A"/></w:rPr></w:pPr>'
        + run("first ")
        + "</w:p>"
        '<w:bookmarkEnd w:id="9"/>'
        '<w:p w14:paraId="00000002"><w:pPr><w:jc w:val="right"/></w:pPr>' + run("second") + "</w:p>"
    )
    revisions.accept_all(package)
    ps = paragraphs(package)
    assert len(ps) == 1
    assert ps[0].para_id == "00000002"
    assert ps[0].p_pr.p_style is None
    assert text_of(ps[0]) == "first second"
    assert len(package.main_document_part.contents.body.content) == 1
    assert any(type(item).__name__ == "CTMarkupRange" for item in ps[0].content), "the marker went into the paragraph"


def test_u2_a_field_straddling_an_edge_resolves_as_runs():
    xml = (
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:del w:id="1" w:author="A"><w:r><w:delText>1</w:delText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "<w:r><w:delText> after</w:delText></w:r></w:del></w:p>"
    )
    package = with_body(xml)
    changes = revisions.tracked_changes(package)
    assert [str(change.kind) for change in changes] == ["run_deletion"]
    rejected = with_body(xml)
    revisions.reject_all(rejected)
    assert 'w:fldCharType="end"' in body_xml(rejected)
    revisions.accept_all(package)
    assert 'w:fldCharType="end"' not in body_xml(package)


def test_c01_an_inserted_field_rejected_goes_whole():
    """C01: Word's form for an inserted field, rejected: begin, instruction, separator, result and end all go.

    CR-003 section 22.3 found Phase F leaving a lone ``separate`` and its result.
    """
    xml = (
        "<w:p>" + run("before ")
        + '<w:ins w:id="1" w:author="A"><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        "<w:r><w:instrText> PAGE </w:instrText></w:r></w:ins>"
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>'
        '<w:ins w:id="2" w:author="A"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:ins></w:p>'
    )
    package = with_body(xml)
    changes = revisions.tracked_changes(package)
    assert [(str(c.kind), c.text) for c in changes] == [("field_insertion", "1")]
    revisions.reject_all(package)
    assert "fldChar" not in body_xml(package) and "PAGE" not in body_xml(package)
    assert text_of(paragraphs(package)[0]) == "before "
    accepted = with_body(xml)
    revisions.accept_all(accepted)
    xml_after = body_xml(accepted)
    assert xml_after.count("<w:fldChar ") == 3 and "<w:ins " not in xml_after


def test_c11_bookmarks_touching_removed_content_go():
    """A bookmark with an end in removed content goes whole (C11); one around removed content goes too (C11b, U15)."""
    xml = (
        "<w:p>" + run("a ") + '<w:bookmarkStart w:id="10" w:name="half"/>' + run("b ")
        + '<w:del w:id="1" w:author="A"><w:r><w:delText>c </w:delText></w:r><w:bookmarkEnd w:id="10"/></w:del>'
        + run("d ")
        + '<w:bookmarkStart w:id="11" w:name="around"/>' + ins(2, "A", "e") + '<w:bookmarkEnd w:id="11"/>'
        + '<w:bookmarkStart w:id="12" w:name="point"/><w:bookmarkEnd w:id="12"/></w:p>'
    )
    package = with_body(xml)
    revisions.accept_all(package)
    out = body_xml(package)
    assert "half" not in out
    assert "around" in out
    assert "point" in out
    rejected = with_body(xml)
    revisions.reject_all(rejected)
    out = body_xml(rejected)
    assert "half" in out
    assert "around" not in out and 'w:id="11"' not in out
    assert "point" in out


def test_u4_a_hyperlinks_relationship_goes_with_its_deleted_row():
    from docx4j_py.openpackaging.parts.namespaces import Namespaces

    package = with_body("<w:p/>")
    rels = package.main_document_part.get_relationships_part(True)
    rel = rels.add_external_relationship(Namespaces.HYPERLINK, "https://www.docx4java.org/")
    package.main_document_part.set_xml(
        f"<w:document {NS}><w:body><w:p>" + run("kept") + "</w:p>"
        '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>'
        "<w:tr><w:tc><w:p>" + run("row one") + "</w:p></w:tc></w:tr>"
        '<w:tr><w:trPr><w:del w:id="1" w:author="A"/></w:trPr><w:tc><w:p>'
        f'<w:hyperlink r:id="{rel.id}"><w:del w:id="2" w:author="A"><w:r><w:delText>link</w:delText></w:r></w:del>'
        "</w:hyperlink></w:p></w:tc></w:tr></w:tbl><w:p/></w:body></w:document>"
    )
    assert rels.get_relationship_by_id(rel.id) is not None
    revisions.accept_all(package)
    assert "row one" in body_xml(package)
    assert rels.get_relationship_by_id(rel.id) is None


def test_u3_a_removed_comment_takes_its_replies_and_entries():
    """A removed comment's entries in commentsExtended, commentsIds and commentsExtensible go, and its reply."""
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "content"))
    from docx4j_py.openpackaging.content_types import ContentTypes
    from docx4j_py.openpackaging.parts.default_xml_part import DefaultXmlPart
    from docx4j_py.openpackaging.parts.namespaces import Namespaces
    from docx4j_py.openpackaging.parts.wml import CommentsExtendedPart, CommentsPart

    package = with_body(
        "<w:p>" + run("before ")
        + '<w:ins w:id="1" w:author="A"><w:commentRangeStart w:id="20"/><w:commentRangeStart w:id="21"/>'
        "<w:r><w:t>commented</w:t></w:r></w:ins>"
        '<w:commentRangeEnd w:id="20"/><w:commentRangeEnd w:id="21"/>'
        '<w:r><w:commentReference w:id="20"/></w:r><w:r><w:commentReference w:id="21"/></w:r>'
        + run(" after")
        + "</w:p>"
    )
    main = package.main_document_part
    w14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    comments = CommentsPart()
    comments.set_xml(
        f"<w:comments {w} {w14}>"
        '<w:comment w:id="20" w:author="A"><w:p w14:paraId="0000000A">' + run("the comment") + "</w:p></w:comment>"
        '<w:comment w:id="21" w:author="B"><w:p w14:paraId="0000000B">' + run("the reply") + "</w:p></w:comment>"
        "</w:comments>"
    )
    main.add_target_part(comments)
    w15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
    extended = CommentsExtendedPart()
    extended.set_xml(
        f"<w15:commentsEx {w15}>"
        '<w15:commentEx w15:paraId="0000000A" w15:done="0"/>'
        '<w15:commentEx w15:paraId="0000000B" w15:paraIdParent="0000000A" w15:done="0"/></w15:commentsEx>'
    )
    main.add_target_part(extended)
    cid = 'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"'
    cex = 'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"'
    for name, content_type, relationship, xml in (
        (
            "/word/commentsIds.xml",
            ContentTypes.WORDPROCESSINGML_COMMENTS_IDS,
            Namespaces.COMMENTS_IDS,
            f"<w16cid:commentsIds {cid}>"
            '<w16cid:commentId w16cid:paraId="0000000A" w16cid:durableId="1000000A"/>'
            '<w16cid:commentId w16cid:paraId="0000000B" w16cid:durableId="1000000B"/></w16cid:commentsIds>',
        ),
        (
            "/word/commentsExtensible.xml",
            ContentTypes.WORDPROCESSINGML_COMMENTS_EXTENSIBLE,
            Namespaces.COMMENTS_EXTENSIBLE,
            f"<w16cex:commentsExtensible {cex}>"
            '<w16cex:commentExtensible w16cex:durableId="1000000A"/>'
            '<w16cex:commentExtensible w16cex:durableId="1000000B"/></w16cex:commentsExtensible>',
        ),
    ):
        part = DefaultXmlPart(name, content_type, relationship)
        part.set_xml(xml)
        main.add_target_part(part)

    revisions.reject_all(package)
    assert text_of(paragraphs(package)[0]) == "before  after"
    assert "comment" not in body_xml(package)
    assert comments.contents.comment == []
    assert extended.contents.comment_ex == []
    assert "commentId " not in main.comments_ids_part.get_xml() and "1000000B" not in main.comments_ids_part.get_xml()
    assert "1000000A" not in main.comments_extensible_part.get_xml()


def test_math_structures():
    """An inserted equation structure: accepted, its control keeps the recorded run properties; rejected, it goes whole."""
    package = fixture("tracked-changes-equations.docx")
    inserted = revisions.tracked_changes(package)[1]
    assert inserted.text.endswith("s-1.")
    assert any(piece.kind == "math_structure_insertion" for piece in inserted.pieces)
    inserted.accept()
    xml = body_xml(package)
    s_sup = xml.index("<m:sSup>")
    ctrl_pr = xml[xml.index("<m:ctrlPr>", s_sup) : xml.index("</m:ctrlPr>", s_sup)]
    assert "w:ins" not in ctrl_pr
    assert "Cambria Math" in ctrl_pr

    rejected = fixture("tracked-changes-equations.docx")
    revisions.tracked_changes(rejected)[1].reject()
    xml = body_xml(rejected)
    first_equation = xml[xml.index("<m:oMathPara>") : xml.index("</m:oMathPara>")]
    assert "<m:sSup>" not in first_equation, "the inserted structure went whole"


# ---------------------------------------------------------------------------
# what this port adds: typed islands, every story, parts left alone
# ---------------------------------------------------------------------------


def _alternate_content_xml(package: object) -> str:
    from docx4j_py.child import is_alternate_content
    from docx4j_py.fragments import to_xml
    from docx4j_py.traversal import iter_nodes
    from docx4j_py.wml import R

    body = package.main_document_part.contents.body
    for node in iter_nodes(body, mce="all"):
        if isinstance(node, R):
            for item in node.content:
                if is_alternate_content(item):
                    return to_xml(item)
    raise AssertionError("no mc:AlternateContent")


def test_a_typed_island_leaves_the_rest_of_its_branch_unchanged():
    """CR-010 section 3.5 and risk 4: only the w:txbxContent is rewritten; everything around it is as it was."""
    import re

    package = fixture("cr043/cr043-tracked.docx")
    before = _alternate_content_xml(package)
    revisions.accept_all(package)
    after = _alternate_content_xml(package)
    boxes = re.compile(r"<w:txbxContent>.*?</w:txbxContent>", re.S)
    # the fragment's namespace declarations are what its content uses: w16du
    # goes with the revisions that carried w16du:dateUtc, and nothing else does
    declarations = re.compile(r'\sxmlns:\w+="[^"]*"')
    assert declarations.sub("", boxes.sub("BOX", before)) == declarations.sub("", boxes.sub("BOX", after)), (
        "the shape, the VML and both branches untouched"
    )
    inside = boxes.findall(after)
    assert len(inside) == 2
    for box in inside:
        assert "<w:ins " not in box and "<w:del " not in box
        assert "C22 inserted." in box and "DELETEME" not in box


def test_a_text_box_with_no_revision_is_never_made_an_island():
    """A document whose text boxes nobody edited keeps them exactly: no island, no rewrite."""
    from docx4j_py.model.revisions import islands

    package = fixture("cr043/cr043-accept-all.docx")  # Word's save: no revision left anywhere
    made: list = []
    original = islands.Island.__init__

    def counting(self, node, holder):  # noqa: ANN001
        made.append(node)
        original(self, node, holder)

    islands.Island.__init__ = counting
    try:
        assert revisions.tracked_changes(package) == []
    finally:
        islands.Island.__init__ = original
    assert made == []


def test_every_story_is_resolved_by_the_package_verbs():
    """CR-010 section 3.3: headers, footers, notes and comments with the body; Word's probe log counts none left."""
    for verb in (revisions.accept_all, revisions.reject_all):
        package = fixture("cr043/cr043-tracked.docx")
        verb(package)
        parts = {str(change.part.part_name) for change in revisions.tracked_changes(package)}
        assert parts == {"/word/document.xml"}, "only Phase B's records are left, in the body"
        saved = round_trip(package)
        for part in saved.main_document_part.header_parts() + saved.main_document_part.footer_parts():
            assert "<w:ins " not in part.get_xml() and "<w:del " not in part.get_xml()
        endnotes = saved.main_document_part.endnotes_part.get_xml()
        assert "<w:ins " not in endnotes and "<w:del " not in endnotes and "rPrChange" not in endnotes


def test_a_bodys_verbs_resolve_its_own_story_and_leave_the_others():
    """``Body.accept_all()`` is the body's story, with Word's side effects; the headers keep theirs."""
    package = fixture("cr043/cr043-tracked.docx")
    package.body.accept_all()
    left = {str(change.part.part_name) for change in package.get_tracked_changes()}
    assert "/word/document.xml" in left, "the table's and the section's records"
    # C06's deleted section break, accepted, took the first section's headers with it (D7)
    assert {"/word/header5.xml", "/word/footer5.xml", "/word/endnotes.xml"} <= left
    assert package.get_part("/word/header2.xml") is None


def test_a_note_goes_with_its_reference():
    """C12: the endnote whose reference was deleted is removed from its part when the deletion is accepted."""
    package = fixture("cr043/cr043-tracked.docx")
    endnotes = package.main_document_part.endnotes_part.contents.endnote
    before = len(endnotes)
    revisions.accept_all(package)
    assert len(package.main_document_part.endnotes_part.contents.endnote) == before - 1


def test_a_picture_part_goes_with_its_last_reference():
    """C04, C05 and D7: Word's saves keep no unreferenced image; accepted, the deleted picture's part goes."""

    def images(package: object) -> int:
        return sum(1 for part in package.parts.values() if (part.content_type or "").startswith("image/"))

    package = fixture("cr043/cr043-tracked.docx")
    before = images(package)
    revisions.accept_all(package)
    after = images(package)
    assert after == before - 1
    with zipfile.ZipFile(io.BytesIO(package.save())) as archive:
        types = archive.read("[Content_Types].xml").decode("utf-8")
        names = set(archive.namelist())
    for part in package.parts.values():
        assert part.part_name.store_name in names or not part.part_name.store_name
    assert types.count("image/") <= types.count("Extension=") + after


def test_listing_does_not_unmarshal_a_part_with_no_revisions():
    """A story part whose bytes hold no revision markup is not read, so it is still written back byte for byte."""
    package = fixture("core-ts/check22/22-deleted.docx")
    main = package.main_document_part
    others = [p for p in (main.footnotes_part, main.endnotes_part, main.comments_part, *main.header_parts()) if p]
    revisions.tracked_changes(package)
    assert all(not part.is_unmarshalled for part in others)


def test_nothing_resolvable_is_left_and_the_result_saves():
    package = fixture("cr043/cr043-tracked.docx")
    revisions.reject_all(package)
    assert unresolved(package) == []
    round_trip(package)
