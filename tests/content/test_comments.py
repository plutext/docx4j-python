"""Comments: the reads, the verbs and the five parts. CR-003 section 3.9, Phase G.

Three documents:

``tests/fixtures/comments.docx``
    docx4j's ``comments-one.docx``: one comment, ``w:comments`` only, no modern
    parts and no ``w14:paraId``.
``tests/fixtures/comments-modern.docx``
    Word-written, with all five parts: a ``w15:commentEx``, a ``w15:person``
    carrying an email, a ``w16cid:durableId`` and a ``w16cex:dateUtc``.
``conftest.threaded_package()``
    built here, because no document in any of the three checkouts carries a
    **reply** thread (CR-003 section 15).
"""

from __future__ import annotations

import zipfile

import pytest
from conftest import (
    FIXTURES,
    fixture,
    part_bytes,
    reloaded,
    sample,
    threaded_package,
)

from docx4j_py.model.content import Author
from docx4j_py.model.content.comments import (
    COMMENT_REFERENCE_STYLE,
    COMMENT_TEXT_STYLE,
    DEFAULT_AUTHOR,
    initials_of,
    markers_of,
)
from docx4j_py.model.content.errors import ContentError, SpanError, StyleError

#: The five parts, by their zip entry names.
COMMENT_PARTS = (
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/people.xml",
    "word/commentsExtensible.xml",
)


def source_bytes(name: str) -> dict[str, bytes]:
    """Every entry of a fixture's zip, by name."""
    with zipfile.ZipFile(FIXTURES / name) as archive:
        return {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}


# ---------------------------------------------------------------------------
# reading
# ---------------------------------------------------------------------------


def test_one_comment_read_from_a_document_word_wrote():
    comments = fixture("comments.docx").body.get_comments()
    assert len(comments) == 1
    comment = comments[0]
    assert comment.id == 0
    assert comment.author_name == "jharrop2"
    assert comment.initials == "j"
    assert comment.content == "One comment"
    assert comment.creation_date is not None
    assert comment.creation_date.year == 2011
    assert comment.resolved is False
    assert comment.replies == []
    assert comment.parent is None
    # comments-one.docx predates w14:paraId, so there is nothing to key on
    assert comment.para_id is None
    assert comment.author_email == ""


def test_the_commented_range_is_the_text_the_markers_surround():
    comment = fixture("comments.docx").body.get_comments()[0]
    ranges = comment.get_range()
    assert len(ranges) == 1
    assert ranges[0].text == "paragraph"


def test_the_modern_parts_give_the_email_the_para_id_and_the_done_flag():
    comment = fixture("comments-modern.docx").body.get_comments()[0]
    assert comment.author_name == "Buck Cronk"
    assert comment.initials == "BC"
    # w15:presenceInfo/@w15:userId is "S::buck.cronk@oracle.com::<guid>"
    assert comment.author_email == "buck.cronk@oracle.com"
    assert comment.para_id == "36593EBD"
    assert comment.resolved is False


def test_a_thread_nests_its_replies_under_their_parent():
    comments = threaded_package().body.get_comments()
    assert [c.id for c in comments] == [0, 2]
    assert [r.id for r in comments[0].replies] == [1]
    reply = comments[0].replies[0]
    assert reply.parent is comments[0]
    assert reply.author_name == "Bob Bones"
    assert reply.author_email == "bob@example.com"
    assert reply.content == "Yes, checked."
    assert comments[1].resolved is True
    assert comments[0].resolved is False


def test_get_comments_on_a_paragraph_and_on_a_range():
    package = threaded_package()
    body = package.body
    first, second = body.paragraphs
    assert [c.id for c in first.get_comments()] == [0]
    assert [c.id for c in second.get_comments()] == [2]
    # "Alpha " is in front of the markers, "beta" is inside them
    assert first.get_range().text == "Alpha beta gamma"
    assert [c.id for c in first.search("beta")[0].get_comments()] == [0]
    assert first.search("gamma")[0].get_comments() == []


def test_a_comment_is_a_body_of_its_own():
    comment = threaded_package().body.get_comments()[0]
    assert comment.comment_body.text == "Is this right?"
    assert [p.style_id for p in comment.paragraphs] == [COMMENT_TEXT_STYLE]
    assert comment.comment_body.to_markdown() == "Is this right?"


def test_to_dict_carries_the_thread():
    comment = threaded_package().body.get_comments()[0]
    out = comment.to_dict()
    assert out["id"] == 0
    assert out["author_name"] == "Ada Lovelace"
    assert out["author_email"] == "ada@example.com"
    assert out["address"] == "w14:0A0A0A0A"
    assert out["replies"][0]["content"] == "Yes, checked."
    assert "resolved" not in out
    assert threaded_package().body.get_comments()[1].to_dict()["resolved"] is True


def test_the_repr_names_the_id_the_address_and_the_author():
    comment = threaded_package().body.get_comments()[0]
    assert repr(comment) == "<Comment 0 w14:0A0A0A0A 'Ada Lovelace' 'Is this right?'>"
    assert str(comment) == "Is this right?"


def test_the_outline_count_agrees_with_the_comments_read_back():
    package = threaded_package()
    comments = package.body.get_comments()
    total = sum(1 + len(c.replies) for c in comments)
    assert package.outline().stats.comments == total == 3


def test_a_block_level_marker_gets_no_range():
    """CR-003 section 4: only markers inside a paragraph are mapped to offsets."""
    package = threaded_package()
    body = package.body
    body.insert_xml('<w:commentRangeStart w:id="9"/>', location="End")
    body.insert_xml('<w:commentRangeEnd w:id="9"/>', location="End")
    ids = [m.id for m in markers_of(body.container)]
    assert ids.count(9) == 2
    from docx4j_py.model.content.comments import Comment

    parts_holder = body.get_comments()[0].parts
    ghost = Comment(parts_holder.comments.comment[0], parts_holder, body)
    object.__setattr__(ghost.element, "id", 9)
    assert ghost.get_range() == []


# ---------------------------------------------------------------------------
# what a read pays for (CR-003 section 3.9: three parts, not five)
# ---------------------------------------------------------------------------


def test_reading_comments_unmarshals_three_parts_and_leaves_two_alone():
    package = fixture("comments-modern.docx")
    package.body.get_comments()
    main = package.main_document_part
    assert main.comments_part.is_unmarshalled
    assert main.comments_extended_part.is_unmarshalled
    assert main.people_part.is_unmarshalled
    assert not main.comments_ids_part.is_parsed
    assert not main.comments_extensible_part.is_parsed


def test_a_document_whose_comments_are_not_read_keeps_all_five_byte_for_byte():
    package = fixture("comments-modern.docx")
    assert package.body.text  # the body is read; the comments are not
    saved = part_bytes(package.save())
    source = source_bytes("comments-modern.docx")
    for name in COMMENT_PARTS:
        assert saved[name] == source[name], name


def test_reading_comments_leaves_the_two_side_parts_byte_for_byte():
    package = fixture("comments-modern.docx")
    package.body.get_comments()
    saved = part_bytes(package.save())
    source = source_bytes("comments-modern.docx")
    for name in ("word/commentsIds.xml", "word/commentsExtensible.xml"):
        assert saved[name] == source[name], name


def test_a_document_that_has_the_comment_styles_keeps_styles_xml_byte_for_byte():
    package = fixture("comments-modern.docx")
    package.body.search("Word")
    paragraph = next(p for p in package.body.iter_paragraphs() if p.text.strip())
    paragraph.insert_comment("A second comment.")
    assert "/word/styles.xml" not in package.last_change.parts_touched
    saved = part_bytes(package.save())
    assert saved["word/styles.xml"] == source_bytes("comments-modern.docx")["word/styles.xml"]


# ---------------------------------------------------------------------------
# the author
# ---------------------------------------------------------------------------


def test_the_default_author_is_the_application_name(new_package):
    assert new_package.author == DEFAULT_AUTHOR
    assert new_package.author.name == "docx4j-python"


def test_the_author_can_be_set_as_an_object_or_as_a_name(new_package):
    new_package.author = Author("Ada Lovelace", initials="AL", email="ada@example.com")
    assert new_package.author.email == "ada@example.com"
    new_package.author = "Bob"
    assert new_package.author == Author("Bob")
    with pytest.raises(ContentError) as raised:
        new_package.author = 7
    assert raised.value.code == "author.invalid"


def test_initials_are_derived_from_the_name_when_none_are_given():
    assert initials_of(Author("Ada Lovelace")) == "AL"
    assert initials_of(Author("Ada Lovelace", initials="a.l.")) == "a.l."
    assert initials_of(Author("docx4j-python")) == "D"


# ---------------------------------------------------------------------------
# inserting
# ---------------------------------------------------------------------------


def commented(text: str = "Changed because the source says 'red'.", find: str = "first"):
    """``2010-sample1.docx`` with a comment on `find`; it has no comment parts."""
    package = sample("2010-sample1.docx")
    package.id_seed = 20260917
    package.author = Author("Claude", initials="C", email="claude@example.com")
    comment = package.body.search(find)[0].insert_comment(text)
    return package, comment


def test_inserting_a_comment_creates_the_four_parts_and_the_two_styles():
    package, comment = commented()
    main = package.main_document_part
    assert main.comments_part is not None
    assert main.comments_extended_part is not None
    assert main.comments_ids_part is not None
    assert main.people_part is not None
    # w16cex is never created (CR-003 section 4)
    assert main.comments_extensible_part is None
    from docx4j_py.model.content.styles import style_ids_of

    defined = style_ids_of(package.style_definitions_part)
    assert COMMENT_TEXT_STYLE in defined
    assert COMMENT_REFERENCE_STYLE in defined
    assert comment.id == 0
    assert comment.author_name == "Claude"
    assert comment.initials == "C"


def test_the_markers_go_around_the_span_and_the_reference_run_follows():
    package, comment = commented()
    xml = package.main_document_part.get_xml()
    assert '<w:commentRangeStart w:id="0"/>' in xml
    assert '<w:commentRangeEnd w:id="0"/>' in xml
    assert '<w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/>' in xml
    assert comment.get_range()[0].text == "first"


def test_the_change_report_names_every_part_the_insert_touched():
    package, _comment = commented()
    report = package.last_change
    assert report.operation == "insert_comment"
    assert report.text_after == "Changed because the source says 'red'."
    assert set(report.parts_touched) == {
        "/word/document.xml",
        "/word/styles.xml",
        "/word/comments.xml",
        "/word/commentsExtended.xml",
        "/word/commentsIds.xml",
        "/word/people.xml",
    }


def test_the_side_part_entries_are_written_and_keyed_by_the_para_id():
    package, comment = commented()
    para_id = comment.para_id
    assert para_id and len(para_id) == 8
    assert f'w15:paraId="{para_id}"' in package.main_document_part.comments_extended_part.get_xml()
    ids_xml = package.main_document_part.comments_ids_part.get_xml()
    assert f'w16cid:paraId="{para_id}"' in ids_xml
    assert 'w16cid:durableId="' in ids_xml
    people = package.main_document_part.people_part.get_xml()
    assert 'w15:author="Claude"' in people
    assert 'w15:userId="claude@example.com"' in people


def test_a_comment_survives_a_save_and_a_reload():
    package, _comment = commented()
    back = reloaded(package).body.get_comments()
    assert len(back) == 1
    assert back[0].content == "Changed because the source says 'red'."
    assert back[0].author_email == "claude@example.com"
    assert back[0].get_range()[0].text == "first"


def test_a_comment_on_a_whole_paragraph():
    package = sample("2010-sample1.docx")
    package.id_seed = 20260917
    paragraph = package.body.paragraphs[0]
    comment = paragraph.insert_comment("The whole paragraph.")
    assert comment.get_range()[0].text == paragraph.text
    assert [c.id for c in paragraph.get_comments()] == [comment.id]


def test_an_empty_range_gets_a_reference_run_only():
    package = sample("2010-sample1.docx")
    package.id_seed = 20260917
    span = package.body.paragraphs[0].get_range("End")
    comment = span.insert_comment("At the end.")
    xml = package.main_document_part.get_xml()
    assert "commentRangeStart" not in xml
    assert '<w:commentReference w:id="0"/>' in xml
    assert comment.get_range()[0].text == ""


def test_the_next_comment_id_is_above_the_highest_in_the_document():
    package = fixture("comments.docx")
    package.id_seed = 20260917
    comment = package.body.paragraphs[0].insert_comment("Second.")
    assert comment.id == 1
    # document order, not id order: the new comment covers the whole paragraph,
    # so its w:commentRangeStart is in front of the one that was already there
    assert [c.id for c in package.body.get_comments()] == [1, 0]


def test_a_span_that_crosses_a_run_holder_is_refused_with_where_to_split():
    package = fixture("hyperlink.docx")
    paragraph = package.body.paragraphs[0]
    span = paragraph.search("a hyperlink")[0]
    with pytest.raises(SpanError) as raised:
        span.insert_comment("crosses the hyperlink")
    assert raised.value.code == "range.crosses_holder"
    assert "insert_comment" in raised.value.message
    assert "split the span at" in raised.value.hint


def test_the_markers_are_hoisted_out_of_a_tracked_insertion():
    """CR-003 section 4: a comment is not a revision."""
    package = sample("sample-docx.docx")
    package.id_seed = 20260917
    span = package.body.search("An insertion")[0]
    comment = span.insert_comment("about the insertion")
    xml = span.paragraph.get_xml()
    start = xml.index(f'<w:commentRangeStart w:id="{comment.id}"/>')
    assert xml.index("<w:ins ") > start, "the start marker must be in front of the w:ins"
    assert '<w:ins w:id="1"' in xml
    assert "</w:ins><w:commentRangeEnd" in xml
    assert comment.get_range()[0].text == "An insertion"


# ---------------------------------------------------------------------------
# replying, resolving, editing and deleting
# ---------------------------------------------------------------------------


def test_a_reply_nests_inside_its_parents_markers():
    package, comment = commented()
    reply = comment.reply("Agreed.")
    assert reply.id == 1
    assert comment.replies == [reply]
    assert reply.parent is comment
    extended = package.main_document_part.comments_extended_part.get_xml()
    assert f'w15:paraIdParent="{comment.para_id}"' in extended
    xml = package.main_document_part.get_xml()
    assert '<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>' in xml
    assert '<w:commentRangeEnd w:id="0"/><w:commentRangeEnd w:id="1"/>' in xml
    back = reloaded(package).body.get_comments()
    assert [c.id for c in back] == [0]
    assert [r.content for r in back[0].replies] == ["Agreed."]


def test_resolving_a_thread_writes_w15_done():
    package, comment = commented()
    assert comment.resolved is False
    comment.resolved = True
    assert comment.resolved is True
    assert 'w15:done="1"' in package.main_document_part.comments_extended_part.get_xml()
    assert package.last_change.operation == "resolve_comment"
    assert reloaded(package).body.get_comments()[0].resolved is True
    comment.resolved = False
    assert package.last_change.operation == "reopen_comment"
    assert comment.resolved is False


def test_setting_the_content_replaces_the_paragraphs_and_keeps_the_para_id():
    package, comment = commented()
    para_id = comment.para_id
    comment.content = "Two\nlines."
    assert comment.content == "Two\nlines."
    assert comment.para_id == para_id
    assert len(comment.paragraphs) == 2
    assert reloaded(package).body.get_comments()[0].content == "Two\nlines."


def test_deleting_a_comment_removes_its_replies_and_every_entry():
    package, comment = commented()
    reply = comment.reply("Agreed.")
    para_ids = {comment.para_id, reply.para_id}
    comment.delete()
    main = package.main_document_part
    assert main.comments_part.contents.comment == []
    assert main.comments_extended_part.contents.comment_ex == []
    ids_xml = main.comments_ids_part.get_xml()
    for para_id in para_ids:
        assert para_id not in ids_xml
    xml = main.get_xml()
    assert "commentRangeStart" not in xml
    assert "commentReference" not in xml
    assert package.body.get_comments() == []
    assert reloaded(package).body.get_comments() == []


def test_deleting_one_comment_of_two_leaves_the_other_alone():
    package = threaded_package()
    comments = package.body.get_comments()
    comments[0].delete()
    left = package.body.get_comments()
    assert [c.id for c in left] == [2]
    assert left[0].content == "Done with this one."
    assert 'w16cid:durableId="1C1C1C1C"' in package.main_document_part.comments_ids_part.get_xml()
    assert 'w16cex:durableId="1C1C1C1C"' in (
        package.main_document_part.comments_extensible_part.get_xml()
    )
    assert "1A1A1A1A" not in package.main_document_part.comments_ids_part.get_xml()


def test_a_w16cex_entry_is_kept_in_step_but_the_part_is_never_created():
    package = threaded_package()
    package.author = Author("Ada Lovelace")
    package.body.paragraphs[1].insert_comment("One more.")
    extensible = package.main_document_part.comments_extensible_part.get_xml()
    assert extensible.count("w16cex:commentExtensible") == 4
    # and the author is already in w:people, so no second entry is written
    assert package.main_document_part.people_part.get_xml().count("w15:person ") == 2


def test_replying_to_a_deleted_comment_says_so():
    package, comment = commented()
    comment.delete()
    with pytest.raises(ContentError) as raised:
        comment.reply("too late")
    assert raised.value.code == "comment.gone"
    assert "get_comments()" in raised.value.hint
    assert package.body.get_comments() == []


def test_a_comment_in_a_header_is_accepted_and_anchored_there():
    """Word does not write them; the API accepts them, as the TypeScript engine does."""
    package = sample("Headers.docx")
    package.id_seed = 20260917
    header = next(b for b in package.bodies() if b.prefix.startswith("header"))
    comment = header.paragraphs[0].insert_comment("in the header")
    assert comment.id == 0
    assert [c.id for c in header.get_comments()] == [0]
    # the w:comment itself lives in the document's comments part, as Word has it
    assert package.main_document_part.comments_part is not None
    assert "/word/comments.xml" in package.last_change.parts_touched


def test_the_comment_styles_are_not_added_twice():
    package, comment = commented()
    first = package.style_definitions_part.get_xml().count('w:styleId="CommentText"')
    comment.reply("Agreed.")
    assert package.style_definitions_part.get_xml().count('w:styleId="CommentText"') == first == 1
    assert "/word/styles.xml" not in package.last_change.parts_touched


def test_a_document_with_no_styles_part_still_takes_a_comment():
    package = sample("2010-sample1.docx")
    package.id_seed = 20260917
    from docx4j_py.model.content.comments import ensure_comment_styles

    package.main_document_part.style_definitions_part = None
    assert ensure_comment_styles(package) == ()
    comment = package.body.paragraphs[0].insert_comment("no styles part here")
    assert comment.id == 0


def test_a_style_that_docx4j_does_not_know_is_still_refused():
    """The escape hatch of CR-003 section 14.9 is unchanged by this phase."""
    package = sample("2010-sample1.docx")
    with pytest.raises(StyleError) as raised:
        package.body.paragraphs[0].style = "Not A Real Style"
    assert raised.value.code == "style.not_found"


def test_get_comments_on_a_content_control():
    package = sample("invoice2013.docx")
    controls = package.body.content_controls
    assert controls
    assert controls[0].get_comments() == []
