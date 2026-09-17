"""Change tracking: the markup, Word's rules, accept and reject. CR-003 §3.8, §4.

Phase F. The mode over ``w:trackRevisions``; the revision markup every mutation
writes while it is on; the rules of section 4 one by one; ``TrackedChange``,
``accept_all()`` and ``reject_all()``; the two views; ``replace_text``. Every
mutation is followed by a save, a reload and a read back, as section 7 requires,
and the parts nothing touched are still byte for byte.
"""

from __future__ import annotations

import datetime
import re
import zipfile

import pytest
from conftest import (
    SAMPLES,
    fixture,
    moves_package,
    part_bytes,
    reloaded,
    sample,
)

from docx4j_py import WordprocessingMLPackage, create_package
from docx4j_py.model.content import Author, TrackedChange
from docx4j_py.model.content.errors import SpanError, TrackedChangeError

#: The date every test fixes, so that the markup is reproducible.
WHEN = datetime.datetime(2026, 9, 17, 12, 0, tzinfo=datetime.UTC)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def tracked(text: str = "The quick brown fox jumps over the lazy dog.", *, author: str = "Claude"):
    """A created document with one paragraph, tracking on, author and date fixed."""
    package = create_package()
    package.id_seed = 20260917
    package.author = Author(author, initials=author[0])
    package.tracked_change_date = WHEN
    paragraph = package.body.insert_paragraph(text)
    package.change_tracking_mode = "TrackAll"
    return package, paragraph


def untracked(text: str = "The quick brown fox jumps over the lazy dog."):
    """The same document with the mode off, for comparing the accepted result."""
    package = create_package()
    package.id_seed = 20260917
    return package, package.body.insert_paragraph(text)


# ---------------------------------------------------------------------------
# the mode (CR-003 section 3.8)
# ---------------------------------------------------------------------------


def test_the_mode_is_off_by_default_and_is_written_to_track_revisions():
    package = create_package()
    assert package.change_tracking_mode == "Off"
    assert package.body.change_tracker is None

    package.change_tracking_mode = "TrackAll"
    assert package.body.change_tracker is not None
    assert "<w:trackRevisions/>" in package.document_settings_part.get_xml()

    back = reloaded(package)
    assert back.change_tracking_mode == "TrackAll"
    back.change_tracking_mode = "Off"
    assert "trackRevisions" not in back.document_settings_part.get_xml()
    assert reloaded(back).change_tracking_mode == "Off"


def test_reading_the_mode_leaves_the_settings_part_byte_for_byte():
    """CR-003 section 16: a read goes through lxml, a write unmarshals."""
    package = sample("2010-sample1.docx")
    package.body  # noqa: B018 - unmarshal the main part, as a caller would
    before = part_bytes(package.save())

    assert package.change_tracking_mode == "Off"
    assert package.document_settings_part.is_unmarshalled is False

    after = part_bytes(package.save())
    assert after["word/settings.xml"] == before["word/settings.xml"]


def test_track_mine_only_is_stored_as_track_all():
    """CR-003 section 4: ``w:trackRevisions`` is a flag; a file cannot tell them apart."""
    package = create_package()
    package.change_tracking_mode = "TrackMineOnly"
    assert package.change_tracking_mode == "TrackMineOnly", "the session remembers what it asked"
    assert reloaded(package).change_tracking_mode == "TrackAll", "the file cannot say"


def test_a_document_with_no_settings_part_gets_one_with_its_relationship():
    package = sample("2010-glow-then-AlternateContent.docx")
    assert package.document_settings_part is None

    package.change_tracking_mode = "TrackAll"
    part = package.document_settings_part
    assert part is not None and str(part.part_name) == "/word/settings.xml"

    with zipfile.ZipFile(__import__("io").BytesIO(package.save())) as archive:
        names = set(archive.namelist())
        types = archive.read("[Content_Types].xml").decode("utf-8")
        rels = archive.read("word/_rels/document.xml.rels").decode("utf-8")
    assert "word/settings.xml" in names
    assert "/word/settings.xml" in types
    assert "settings.xml" in rels
    assert reloaded(package).change_tracking_mode == "TrackAll"


def test_setting_the_mode_is_one_change_report_naming_the_settings_part():
    package = sample("2010-sample1.docx")
    package.change_tracking_mode = "TrackAll"

    change = package.last_change
    assert change.operation == "change_tracking_mode"
    assert change.text_after == "TrackAll"
    assert "/word/settings.xml" in change.parts_touched
    assert len(package.changes) == 1

    # turning it off again touches the same part, and nothing else
    package.change_tracking_mode = "Off"
    assert package.last_change.parts_touched == change.parts_touched


def test_an_invalid_mode_and_an_invalid_date_are_refused_with_a_hint():
    package = create_package()
    with pytest.raises(Exception) as bad_mode:
        package.change_tracking_mode = "Sometimes"
    assert bad_mode.value.code == "tracking.mode_invalid"
    with pytest.raises(Exception) as bad_date:
        package.tracked_change_date = "2026-09-17"
    assert bad_date.value.code == "tracking.date_invalid"
    assert "datetime" in bad_date.value.hint


# ---------------------------------------------------------------------------
# who, when and which id (CR-003 section 4)
# ---------------------------------------------------------------------------


def test_the_author_is_the_packages_and_only_the_name_reaches_the_markup():
    package, paragraph = tracked(author="Ada Lovelace")
    package.author = Author("Ada Lovelace", initials="AL", email="ada@example.invalid")
    paragraph.insert_text("New. ", location="Start")

    xml = paragraph.get_xml()
    assert 'w:author="Ada Lovelace"' in xml
    assert "ada@example.invalid" not in xml and 'w:initials' not in xml
    assert package.body.get_tracked_changes()[0].author == "Ada Lovelace"


def test_the_date_is_the_fixed_one_or_now_in_utc_to_the_second():
    package, paragraph = tracked()
    paragraph.insert_text("New. ", location="Start")
    assert 'w:date="2026-09-17T12:00:00Z"' in paragraph.get_xml()
    assert package.body.get_tracked_changes()[0].date == WHEN

    package.tracked_change_date = None
    paragraph.insert_text(" More.", location="End")
    when = package.body.get_tracked_changes()[-1].date
    now = datetime.datetime.now(datetime.UTC)
    assert when.tzinfo is not None and when.microsecond == 0
    assert abs((now - when).total_seconds()) < 60


def test_revision_ids_start_above_the_highest_annotation_id_in_use():
    """CR-003 section 4: the one annotation id space, over the parts unmarshalled."""
    package = sample("sample-docx.docx")
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    package.change_tracking_mode = "TrackAll"
    highest = max(
        int(m) for m in re.findall(r'<w:(?:ins|del|bookmarkStart) w:id="(\d+)"', package.body.get_xml())
    )

    package.body.paragraphs[0].insert_text("New. ", location="Start")
    made = [c for c in package.body.get_tracked_changes() if c.author == "Claude"]
    assert made and made[0].id > highest


def test_the_revision_ids_and_the_comment_ids_are_separate_spaces():
    """CR-003 section 4: ``w:comment/@w:id`` is excluded from the revision scan."""
    package, paragraph = tracked("My first 2010 document.")
    comment = paragraph.search("first")[0].insert_comment("Why?")
    paragraph.insert_text("New. ", location="Start")

    revision = package.body.get_tracked_changes()[0]
    assert comment.id == 0, "the comment space starts at zero"
    assert revision.id >= 1
    # a second comment still counts comments, not revisions
    second = paragraph.search("document")[0].insert_comment("And this?")
    assert second.id == 1


# ---------------------------------------------------------------------------
# Word's rules, not just the markup (CR-003 section 4)
# ---------------------------------------------------------------------------


def test_an_insertion_is_a_w_ins_and_the_accepted_text_is_the_untracked_one():
    package, paragraph = tracked()
    paragraph.insert_text("Hello. ", location="Start")

    assert "<w:ins " in paragraph.get_xml()
    assert paragraph.text == "Hello. The quick brown fox jumps over the lazy dog."
    assert paragraph.get_text(view="original") == "The quick brown fox jumps over the lazy dog."

    _plain, plain_paragraph = untracked()
    plain_paragraph.insert_text("Hello. ", location="Start")
    assert paragraph.text == plain_paragraph.text

    back = reloaded(package).body.paragraphs[0]
    assert back.text == paragraph.text
    assert back.get_text(view="original") == "The quick brown fox jumps over the lazy dog."


def test_a_second_insertion_by_the_same_author_extends_the_w_ins():
    _package, paragraph = tracked()
    paragraph.insert_text("A", location="Start")
    paragraph.insert_text("B", location="Start")

    xml = paragraph.get_xml()
    assert xml.count("<w:ins ") == 1, "extended, not nested"
    assert "<w:t>BA</w:t>" in xml
    assert paragraph.text.startswith("BAThe quick")


def test_another_authors_insertion_gets_a_w_ins_of_its_own():
    package, paragraph = tracked()
    paragraph.insert_text("Mine. ", location="Start")
    package.author = Author("Someone Else")
    paragraph.insert_text("Theirs. ", location="Start")

    xml = paragraph.get_xml()
    assert xml.count("<w:ins ") == 2
    assert [c.author for c in package.body.get_tracked_changes()] == [
        "Someone Else",
        "Claude",
    ]


def test_a_deletion_moves_the_runs_into_a_w_del_with_del_text():
    package, paragraph = tracked()
    paragraph.search("quick ")[0].delete()

    xml = paragraph.get_xml()
    assert "<w:del " in xml and "<w:delText" in xml
    assert paragraph.text == "The brown fox jumps over the lazy dog."
    assert paragraph.get_text(view="original") == "The quick brown fox jumps over the lazy dog."

    back = reloaded(package).body.paragraphs[0]
    assert back.text == "The brown fox jumps over the lazy dog."
    assert back.get_text(view="original") == "The quick brown fox jumps over the lazy dog."


def test_deleted_text_renames_w_t_and_w_instr_text_and_back():
    """CR-003 section 4: ``w:t`` becomes ``w:delText``, ``w:instrText`` ``w:delInstrText``.

    Directly on a run, because a field instruction is not *text*: a span never
    covers one, so through the views the rename fires on the ``w:t`` alone and a
    ``w:instrText`` is split into a run of its own and left as it is (CR-003
    section 16.2).
    """
    from docx4j_py.child import ChildList, link_parents
    from docx4j_py.model.content.tracking import to_deleted_text, to_restored_text
    from docx4j_py.wml import R, el, to_xml

    run = R(content=ChildList([el.t("Text"), el.instrText(" PAGE ")]))
    link_parents(run)

    to_deleted_text(run)
    assert "<w:delText>Text</w:delText>" in to_xml(run)
    assert '<w:delInstrText xml:space="preserve"> PAGE </w:delInstrText>' in to_xml(run)

    to_restored_text(run)
    assert "<w:t>Text</w:t>" in to_xml(run)
    assert '<w:instrText xml:space="preserve"> PAGE </w:instrText>' in to_xml(run)


def test_a_replacement_is_the_del_first_and_the_ins_after_it():
    _package, paragraph = tracked()
    assert paragraph.replace_text("quick", "slow") == 1

    xml = paragraph.get_xml()
    assert xml.index("<w:del ") < xml.index("<w:ins ")
    assert paragraph.text == "The slow brown fox jumps over the lazy dog."
    assert paragraph.get_text(view="original") == "The quick brown fox jumps over the lazy dog."


def test_deleting_text_this_author_inserted_takes_it_back():
    package, paragraph = tracked()
    paragraph.insert_text("XYZ ", location="Start")
    paragraph.search("XYZ ")[0].delete()

    assert "<w:ins " not in paragraph.get_xml(), "the emptied w:ins goes with it"
    assert paragraph.text == "The quick brown fox jumps over the lazy dog."
    assert package.body.get_tracked_changes() == []


def test_editing_another_authors_deleted_text_is_refused_by_name():
    package, paragraph = tracked()
    paragraph.search("quick")[0].delete()
    package.author = Author("Someone Else")

    from docx4j_py.model.content.text_model import runs_of
    from docx4j_py.model.content.tracking import revision_of, tracker_of

    deleted = next(
        run
        for run in runs_of(paragraph.element, view="original")
        if (found := revision_of(run)) is not None and found.kind == "del"
    )
    with pytest.raises(TrackedChangeError) as refused:
        tracker_of(package).assert_editable(revision_of(deleted))
    assert refused.value.code == "tracking.deleted_text"
    assert "Claude" in str(refused.value)
    assert "reject" in refused.value.hint


def test_the_runs_are_split_at_the_spans_boundaries():
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    paragraph = package.body.insert_paragraph("Hello brave new world")
    paragraph.search("brave")[0].font.bold = True
    package.change_tracking_mode = "TrackAll"

    paragraph.search("ave new")[0].delete()

    xml = paragraph.get_xml()
    assert "<w:t>br</w:t>" in xml, "the bold run keeps the part that stays"
    assert "<w:delText>ave</w:delText>" in xml
    assert paragraph.text == "Hello br world"
    assert paragraph.get_text(view="original") == "Hello brave new world"


def test_a_paragraph_inserted_in_the_middle_carries_its_own_mark():
    package, paragraph = tracked()
    package.change_tracking_mode = "Off"
    last = paragraph.insert_paragraph("The last one.", location="After")
    package.change_tracking_mode = "TrackAll"

    made = paragraph.insert_paragraph("A new paragraph.", location="After")

    assert "<w:rPr><w:ins " in made.get_xml().replace("\n", ""), "its own mark"
    assert "<w:t>A new paragraph.</w:t>" in made.get_xml()
    assert [(c.kind, c.type) for c in made.get_tracked_changes()] == [
        ("mark", "Added"),
        ("run", "Added"),
    ]
    assert last.get_tracked_changes() == [], "the paragraph after it is untouched"


def test_a_paragraph_appended_at_the_end_leaves_the_final_mark_alone():
    """CR-003 section 16.10: Word never marks a container's final mark inserted."""
    package, paragraph = tracked()
    made = package.body.insert_paragraph("Appended at the end.")

    assert "<w:rPr><w:ins " not in made.get_xml().replace("\n", ""), "the final mark"
    assert "<w:ins " in made.get_xml(), "but the runs are an insertion"
    assert "<w:rPr><w:ins " in paragraph.get_xml().replace("\n", ""), "the mark before it"
    assert [(c.kind, c.type) for c in package.body.get_tracked_changes()] == [
        ("mark", "Added"),
        ("run", "Added"),
    ]

    plain, _plain_paragraph = untracked()
    plain.body.insert_paragraph("Appended at the end.")
    assert package.body.text == plain.body.text

    back = reloaded(package)
    assert back.body.accept_all() == 2
    assert back.body.text == plain.body.text


def test_rejecting_an_appended_paragraph_joins_it_into_the_one_before():
    package, paragraph = tracked("First.")
    first_id = paragraph.para_id
    package.change_tracking_mode = "Off"
    paragraph.style_id = "Heading1"  # the document's own formatting, not a revision
    package.change_tracking_mode = "TrackAll"
    package.body.insert_paragraph("Appended.")

    assert reloaded(package).body.reject_all() == 2
    back = reloaded(package)
    back.body.reject_all()
    paragraphs = back.body.paragraphs
    assert [p.text for p in paragraphs] == ["First."]
    assert paragraphs[0].para_id == first_id, "the surviving paragraph is the one that was there"
    assert paragraphs[0].style_id == "Heading1", "and it keeps its own properties"
    assert back.body.get_tracked_changes() == [], "and the mark goes with the break"
    assert "w:ins" not in back.body.get_xml()


def test_a_fragment_appended_at_the_end_shifts_every_mark_back_one():
    package, _paragraph = tracked("First.")
    package.body.insert_markdown("## A heading\n\npara one\n\npara two\n")

    marked = [
        p.text
        for p in package.body.paragraphs
        if p.element.p_pr is not None
        and p.element.p_pr.r_pr is not None
        and p.element.p_pr.r_pr.ins is not None
    ]
    assert marked == ["First.", "A heading", "para one"], "never the last one"

    back = reloaded(package)
    assert back.body.reject_all() == 6
    assert [p.text for p in back.body.paragraphs] == ["First."], "no empty husk left behind"
    assert back.body.get_tracked_changes() == []
    assert "w:ins" not in back.body.get_xml()


def test_deleting_a_paragraph_marks_the_mark_and_the_content():
    package, paragraph = tracked()
    package.change_tracking_mode = "Off"
    paragraph.insert_paragraph("Second.", location="After")
    package.change_tracking_mode = "TrackAll"
    paragraph.delete()

    assert [p.text for p in package.body.paragraphs] == ["", "Second."]
    kinds = [(c.kind, c.type) for c in paragraph.get_tracked_changes()]
    assert kinds == [("mark", "Deleted"), ("run", "Deleted")]
    assert paragraph.get_text(view="original") == "The quick brown fox jumps over the lazy dog."

    with pytest.raises(TrackedChangeError) as twice:
        paragraph.delete()
    assert twice.value.code == "tracking.already_deleted"


def test_a_paragraph_this_author_inserted_and_then_deleted_simply_goes():
    package, paragraph = tracked()
    made = paragraph.insert_paragraph("Added then removed.", location="After")
    assert len(package.body.paragraphs) == 2

    made.delete()
    assert [p.text for p in package.body.paragraphs] == [
        "The quick brown fox jumps over the lazy dog."
    ]


# ---------------------------------------------------------------------------
# formatting (w:rPrChange and w:pPrChange)
# ---------------------------------------------------------------------------


def test_a_font_change_records_the_properties_that_were_there_once():
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    paragraph = package.body.insert_paragraph("Formatted.")
    paragraph.font.italic = True
    package.change_tracking_mode = "TrackAll"

    paragraph.font.bold = True
    paragraph.font.underline = "Single"

    xml = paragraph.get_xml()
    assert xml.count("<w:rPrChange ") == 1, "recorded once, on the first write"
    original = re.search(r"<w:rPrChange [^>]*>(.*)</w:rPrChange>", xml).group(1)
    assert "<w:i/>" in original and "<w:b/>" not in original


def test_a_paragraph_property_change_records_a_ct_pprbase_with_no_rpr():
    _package, paragraph = tracked()
    paragraph.alignment = "Centered"
    paragraph.style_id = "Heading1"
    paragraph.left_indent = 36

    xml = paragraph.get_xml()
    assert xml.count("<w:pPrChange ") == 1
    original = re.search(r"<w:pPrChange [^>]*>(.*)</w:pPrChange>", xml).group(1)
    assert "xsi:type" not in original, "CT_PPrBase, as Word writes it"
    assert "<w:rPr" not in original and "<w:sectPr" not in original


def test_a_run_this_author_inserted_takes_formatting_with_no_r_pr_change():
    _package, paragraph = tracked()
    span = paragraph.insert_text("New text. ", location="Start")
    span.font.bold = True

    assert "<w:rPrChange" not in paragraph.get_xml()
    assert [c.kind for c in paragraph.get_tracked_changes()] == ["run"]


# ---------------------------------------------------------------------------
# rows and tables (CR-003 section 14.7)
# ---------------------------------------------------------------------------


def test_row_insertions_and_deletions_go_on_the_tr_pr_and_the_row_stays():
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    table = package.body.insert_table(2, 2, values=[["a", "b"], ["c", "d"]])
    package.change_tracking_mode = "TrackAll"

    table.add_rows(1, values=[["e", "f"]])
    table.delete_rows(0)

    xml = table.get_xml()
    assert "<w:trPr><w:ins " in xml and "<w:trPr><w:del " in xml
    assert table.row_count == 3, "a deleted row stays in the tree until accepted"
    assert [c.type for c in package.body.get_tracked_changes() if c.kind == "row"] == [
        "Deleted",
        "Added",
    ]

    back = reloaded(package)
    assert back.body.tables[0].row_count == 3
    assert back.body.accept_all()
    assert back.body.tables[0].values == [["c", "d"], ["e", "f"]]


def test_a_deleted_row_is_emptied_as_Word_empties_one():
    """CR-003 section 16.10: the row mark alone is what Word shows pink, not struck."""
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    table = package.body.insert_table(2, 2, values=[["a", "b"], ["c", "d"]])
    package.change_tracking_mode = "TrackAll"

    table.delete_rows(0)

    row = table.get_xml().partition("</w:tr>")[0]
    assert "<w:trPr><w:del " in row, "the row mark"
    assert row.count("<w:rPr><w:del ") == 2, "every cell paragraph's mark"
    assert row.count("<w:delText>") == 2 and "<w:t>" not in row, "and every run's text"
    assert table.row_count == 2, "the row itself stays until the change is accepted"

    changes = package.body.get_tracked_changes()
    assert [(c.kind, c.type, c.text) for c in changes] == [("row", "Deleted", "a\nb")]

    back = reloaded(package)
    assert back.body.accept_all() == 1
    assert back.body.tables[0].values == [["c", "d"]]

    again = reloaded(package)
    assert again.body.reject_all() == 1
    assert again.body.tables[0].values == [["a", "b"], ["c", "d"]]
    assert "w:del" not in again.body.tables[0].get_xml(), "and nothing of the deletion is left"


def test_a_row_this_author_inserted_and_then_deleted_is_taken_back():
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    table = package.body.insert_table(1, 1, values=[["a"]])
    package.change_tracking_mode = "TrackAll"

    table.add_rows(1, values=[["b"]])
    table.delete_rows(1)

    assert table.row_count == 1 and table.values == [["a"]]
    assert package.body.get_tracked_changes() == []


def test_a_row_delete_and_a_table_delete_mark_rather_than_remove():
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    table = package.body.insert_table(2, 1, values=[["a"], ["b"]])
    package.change_tracking_mode = "TrackAll"

    table.rows[0].delete()
    assert table.row_count == 2
    table.delete()
    assert package.body.tables, "the table is still there"
    assert len([c for c in package.body.get_tracked_changes() if c.kind == "row"]) == 2

    package.body.accept_all()
    assert package.body.tables == [], "every row accepted away takes the table with it"


def test_an_inserted_table_marks_every_row_and_every_paragraph():
    package, _paragraph = tracked()
    table = package.body.insert_table(2, 2, values=[["a", "b"], ["c", "d"]])

    xml = table.get_xml()
    assert xml.count("<w:trPr><w:ins ") == 2, "both rows"
    assert xml.count("<w:rPr><w:ins ") == 4, "and every cell paragraph's mark"
    assert xml.count("<w:ins ") == 10, "and a w:ins around every cell's runs"

    changes = package.body.get_tracked_changes()
    assert [(c.kind, c.type) for c in changes] == [("row", "Added"), ("row", "Added")], (
        "one change per row: the cell-level insertions are folded into it (section 16.10)"
    )
    assert package.body.reject_all() == 2
    assert package.body.tables == [], "and the emptied w:tbl goes with the rows"


# ---------------------------------------------------------------------------
# the other verbs that inherit the tracked primitives
# ---------------------------------------------------------------------------


def test_the_text_setter_and_body_clear_are_tracked():
    package, paragraph = tracked()
    paragraph.text = "Replaced outright."
    assert paragraph.text == "Replaced outright."
    assert paragraph.get_text(view="original") == "The quick brown fox jumps over the lazy dog."

    package.body.clear()
    assert package.body.paragraphs, "nothing is removed while tracking is on"
    assert all(c.type == "Deleted" for c in package.body.get_tracked_changes() if c.kind == "mark")


def test_an_inserted_picture_is_a_run_insertion():
    png = _sample_png()
    package, paragraph = tracked()
    paragraph.insert_inline_picture(png, width=60, alt_text_description="A picture")

    xml = paragraph.get_xml()
    assert "<w:ins " in xml and xml.index("<w:ins ") < xml.index("<w:drawing>")
    assert len(paragraph.inline_pictures) == 1

    package.body.reject_all()
    assert paragraph.inline_pictures == []


def test_a_content_control_inherits_tracking_and_its_delete_does_not():
    package = sample("invoice2013.docx")
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    package.change_tracking_mode = "TrackAll"
    control = package.body.content_controls[0]

    control.insert_text("Tracked through the paragraph primitives")
    assert any(c.kind == "run" for c in package.body.get_tracked_changes())

    before = len(package.body.content_controls)
    control.delete()
    assert len(package.body.content_controls) == before - 1, "deleting a control is not tracked"


def test_markdown_inserted_while_tracking_is_on_is_one_report_and_all_insertions():
    package, _paragraph = tracked()
    package.body.insert_markdown("## A heading\n\nAnd a paragraph.\n")

    assert package.last_change.operation == "insert_markdown"
    added = [c for c in package.body.get_tracked_changes() if c.type == "Added"]
    assert len(added) >= 4, "a mark and a run for each of the two blocks"
    package.body.reject_all()
    assert [p.text for p in package.body.paragraphs] == [
        "The quick brown fox jumps over the lazy dog."
    ]


def test_a_comment_made_while_tracking_is_on_is_a_comment_not_an_insertion():
    """CR-003 section 15.6: the marker hoist must not be undone."""
    package, paragraph = tracked("Base text.")
    paragraph.insert_text("INSERTED ", location="Start")
    paragraph.search("INSERTED")[0].insert_comment("On the insertion.")

    xml = paragraph.get_xml()
    assert xml.index("<w:commentRangeStart") < xml.index("<w:ins "), "hoisted out of the w:ins"
    assert xml.index("</w:ins>") < xml.index("<w:commentRangeEnd")
    reference = xml[xml.index("<w:commentReference") - 200 : xml.index("<w:commentReference")]
    assert "<w:ins " not in reference, "the reference run is not itself an insertion"
    assert all(c.kind != "run" or c.author == "Claude" for c in package.body.get_tracked_changes())

    package.body.accept_all()
    assert [c.content for c in package.body.get_comments()] == ["On the insertion."]
    assert package.body.paragraphs[0].text == "INSERTED Base text."


def test_a_comment_inside_an_insertion_is_hoisted_out_of_it(tmp_path):
    """CR-003 section 16.9: the defect that made a reject take the comment with it.

    A comment on text that lies **wholly inside** this author's own ``w:ins``
    --- which is what commenting on a replacement gives --- had its markers and
    its reference run written *inside* the ``w:ins``, because the replacement's
    ``w:del`` and ``w:ins`` carried a stale ``parent`` and the hoist of section
    15.6 could not find the list they were in.
    """
    package = sample("2010-sample1.docx")
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    package.change_tracking_mode = "TrackAll"
    package.body.replace_text("document", "report")
    package.body.range_of(package.find("report")[0]).insert_comment("why")

    xml = package.body.paragraphs[0].get_xml()
    assert "<w:commentRangeStart" in xml
    assert xml.index("<w:commentRangeStart") < xml.index("<w:ins "), "the start is a sibling"
    assert xml.index("</w:ins>") < xml.index("<w:commentRangeEnd"), "and so is the end"
    inside = xml[xml.index("<w:ins ") : xml.index("</w:ins>")]
    assert "commentRange" not in inside and "commentReference" not in inside

    # the markup the hoist rests on: the w:del and the w:ins of a replacement
    # are children of the paragraph, not of one another
    paragraph = package.body.paragraphs[0]
    from docx4j_py.traversal import element_name

    for item in paragraph.element.content:
        if element_name(item) in (f"{{{_W}}}ins", f"{{{_W}}}del"):
            assert item.parent is paragraph.element


def test_a_comment_on_an_insertion_survives_accepting_and_rejecting_it():
    def commented():
        package = sample("2010-sample1.docx")
        package.author = Author("Claude")
        package.tracked_change_date = WHEN
        package.change_tracking_mode = "TrackAll"
        package.body.replace_text("document", "report")
        package.body.range_of(package.find("report")[0]).insert_comment("why")
        return reloaded(package)

    accepted = commented()
    assert accepted.body.accept_all() == 2
    comments = accepted.body.get_comments()
    assert [c.content for c in comments] == ["why"]
    assert [r.text for r in comments[0].get_range()] == ["report"]

    rejected = commented()
    assert rejected.body.reject_all() == 2
    comments = rejected.body.get_comments()
    assert [c.content for c in comments] == ["why"], "the insertion did not take it with it"
    assert [r.text for r in comments[0].get_range()] == [""], "an empty range, where it was"


def test_a_comment_on_a_paragraph_inserted_while_tracking_is_on_survives_both():
    def commented():
        package = sample("2010-sample1.docx")
        package.author = Author("Claude")
        package.tracked_change_date = WHEN
        package.change_tracking_mode = "TrackAll"
        paragraph = package.body.insert_paragraph("A whole new paragraph.")
        paragraph.insert_comment("about the new paragraph")
        return package

    xml = commented().body.paragraphs[-1].get_xml()
    assert xml.index("<w:commentRangeStart") < xml.index("<w:ins ")
    inside = xml[xml.index("<w:ins ") : xml.index("</w:ins>")]
    assert "commentRange" not in inside and "commentReference" not in inside

    accepted = reloaded(commented())
    accepted.body.accept_all()
    assert [c.content for c in accepted.body.get_comments()] == ["about the new paragraph"]

    rejected = reloaded(commented())
    rejected.body.reject_all()
    assert [c.content for c in rejected.body.get_comments()] == ["about the new paragraph"]


# ---------------------------------------------------------------------------
# TrackedChange: reading, accepting, rejecting
# ---------------------------------------------------------------------------


def test_the_tracked_changes_of_a_word_written_document():
    package = fixture("tracked-pprchange.docx")
    changes = package.body.get_tracked_changes()

    assert [c.type for c in changes] == [
        "Formatted",
        "Added",
        "Added",
        "Added",
        "Added",
    ]
    assert {c.author for c in changes} == {"jharrop"}
    assert changes[0].kind == "paragraph_properties"
    assert changes[1].kind == "mark"
    assert changes[2].text.startswith("This document intentionally contains revisions")
    assert changes[0].date == datetime.datetime(2012, 6, 27, 21, 15, tzinfo=datetime.UTC)
    assert isinstance(changes[0], TrackedChange)
    assert changes[0].to_dict()["address"] == "body/0"


def test_the_moves_the_formatting_change_and_the_rows_of_the_built_document():
    package = moves_package()
    changes = package.body.get_tracked_changes()

    assert [(c.kind, c.type) for c in changes] == [
        ("run", "Deleted"),
        ("run", "Added"),
        ("run_properties", "Formatted"),
        ("row", "Added"),
        ("row", "Deleted"),
    ], "the inserted row's own w:ins is folded into the row change (section 16.10)"
    assert changes[0].text == "Moved sentence.", "a w:moveFrom reads its w:delText"
    assert changes[1].text == "Moved sentence."
    assert package.body.get_text().splitlines()[0] == " Stays."
    assert package.body.get_text(view="original").splitlines()[0] == "Moved sentence. Stays."


def test_accepting_each_kind_does_what_docx4j_does():
    package = moves_package()
    assert package.body.accept_all() == 5

    text = [p.text for p in package.body.paragraphs]
    assert text == [" Stays.", "Moved sentence.", "Was italic, now bold.", "New row"]
    xml = package.body.get_xml()
    assert "moveFrom" not in xml and "moveTo" not in xml
    assert "rPrChange" not in xml, "accepting drops the formatting revisions too"
    assert "<w:b/>" in xml, "the current properties are what survives"


def test_rejecting_each_kind_puts_back_what_was_there():
    package = moves_package()
    assert package.body.reject_all() == 5

    text = [p.text for p in package.body.paragraphs]
    assert text == ["Moved sentence. Stays.", "", "Was italic, now bold.", "Old row"]
    xml = package.body.get_xml()
    assert "<w:i/>" in xml and "<w:b/>" not in xml, "the recorded original is put back"
    assert "rPrChange" not in xml


def test_accepting_a_deleted_paragraph_mark_joins_with_the_next():
    package, paragraph = tracked("First.")
    second = paragraph.insert_paragraph("Second.", location="After")
    second.style_id = "Heading1"
    package.change_tracking_mode = "Off"
    package.change_tracking_mode = "TrackAll"
    first_id = paragraph.para_id
    paragraph.delete()

    assert package.body.accept_all()
    paragraphs = package.body.paragraphs
    assert [p.text for p in paragraphs] == ["Second."]
    assert paragraphs[0].para_id == first_id, "the first one's paraId survives"
    assert paragraphs[0].style_id == "Heading1", "and the second one's properties"


def test_rejecting_an_inserted_paragraph_at_the_end_joins_with_the_one_before():
    package, _paragraph = tracked("First.")
    package.body.insert_paragraph("Added at the end.")

    assert package.body.reject_all()
    assert [p.text for p in package.body.paragraphs] == ["First."]


def test_rejecting_both_mark_forms_of_a_loaded_document_leaves_the_source():
    """CR-003 section 16.11: the mark goes with the break it recorded.

    The two forms in one document --- a paragraph inserted in the **middle**,
    which carries its own mark, and one **appended**, whose mark sits on the
    paragraph before it --- rejected over a document that was loaded rather than
    created, both through ``reject_all()`` and one change at a time in reverse.
    """
    source = [p.text for p in sample("2010-sample1.docx").body.paragraphs]

    def edited():
        package = sample("2010-sample1.docx")
        package.author = Author("A")
        package.tracked_change_date = WHEN
        package.change_tracking_mode = "TrackAll"
        package.body.paragraphs[0].insert_paragraph("mid", location="After")
        package.body.insert_paragraph("end")
        return package

    package = edited()
    assert [c.kind for c in package.get_tracked_changes()] == ["mark", "run", "mark", "run"]

    assert package.body.reject_all() == 4, "what it actually undid"
    assert package.body.get_tracked_changes() == [], "nothing of the revisions is left"
    assert [p.text for p in package.body.paragraphs] == source
    xml = package.body.get_xml()
    assert "w:ins" not in xml and "w:del" not in xml

    one_at_a_time = edited()
    for change in reversed(one_at_a_time.get_tracked_changes()):
        change.reject()
    assert one_at_a_time.get_tracked_changes() == []
    assert [p.text for p in one_at_a_time.body.paragraphs] == source

    accepted = edited()
    assert accepted.body.accept_all() == 4
    assert accepted.get_tracked_changes() == []
    assert [p.text for p in accepted.body.paragraphs] == [
        "My first 2010 document.",
        "mid",
        "",
        "",
        "",
        "end",
    ]


def test_accepting_a_change_that_is_gone_says_so():
    package, paragraph = tracked()
    paragraph.insert_text("New. ", location="Start")
    change = package.body.get_tracked_changes()[0]
    change.accept()

    with pytest.raises(TrackedChangeError) as gone:
        change.accept()
    assert gone.value.code == "tracking.gone"
    assert "get_tracked_changes" in gone.value.hint


def test_get_tracked_changes_on_a_range_covers_the_span_only():
    _package, paragraph = tracked()
    paragraph.replace_text("quick", "slow")
    paragraph.insert_text(" Tail.", location="End")

    whole = paragraph.get_tracked_changes()
    assert len(whole) == 3
    span = paragraph.search("slow")[0].get_tracked_changes()
    assert [c.type for c in span] == ["Deleted", "Added"]


def test_a_tracked_change_has_a_range_a_repr_and_a_to_dict():
    package, paragraph = tracked()
    paragraph.replace_text("quick", "slow")
    deletion, insertion = package.body.get_tracked_changes()

    assert insertion.get_range().text == "slow"
    assert len(deletion.get_range()) == 0, "a deletion's range in the accepted view is collapsed"
    assert repr(insertion).startswith("<TrackedChange Added 'Claude'")
    assert insertion.to_dict() == {
        "type": "Added",
        "author": "Claude",
        "date": "2026-09-17T12:00:00+00:00",
        "text": "slow",
        "id": insertion.id,
        "kind": "run",
        "address": paragraph.address,
    }
    assert insertion == package.body.get_tracked_changes()[1]
    assert len({insertion, package.body.get_tracked_changes()[1]}) == 1


# ---------------------------------------------------------------------------
# the two views, the markdown view and the stats
# ---------------------------------------------------------------------------


def test_the_original_view_is_offered_on_a_body_and_a_paragraph_and_refused_on_a_range():
    package = sample("sample-docx.docx")
    body = package.body
    assert "An insertion" in body.get_text()
    assert "An insertion" not in body.get_text(view="original")
    assert "A deletion" in body.get_text(view="original")

    span = body.paragraphs[0].get_range()
    with pytest.raises(SpanError) as refused:
        span.get_text(view="original")
    assert refused.value.code == "range.no_original_view"
    assert "paragraph" in refused.value.hint


def test_the_markup_view_agrees_with_the_tracked_change_view():
    package = sample("sample-docx.docx")
    changes = package.body.get_tracked_changes()
    markup = package.body.to_markdown(view="markup")

    inserted = [c.text for c in changes if c.type == "Added"]
    deleted = [c.text for c in changes if c.type == "Deleted"]
    assert re.findall(r"\{\+\+(.*?)\+\+\}", markup) == inserted
    assert re.findall(r"\{--(.*?)--\}", markup) == deleted


def test_the_stats_and_describe_agree_with_the_views():
    package = sample("sample-docx.docx")
    assert package.outline().stats.tracked_changes == len(package.body.get_tracked_changes())
    assert package.describe().authors["revisions"] == ["Jason Harrop"]
    assert package.describe().tracking_on is False

    package.change_tracking_mode = "TrackAll"
    assert package.describe().tracking_on is True


# ---------------------------------------------------------------------------
# replace_text (CR-003 section 3.8)
# ---------------------------------------------------------------------------


def test_replace_text_counts_and_reports_the_pair_tracked_and_not():
    plain = sample("2010-sample1.docx")
    count = plain.body.replace_text("document", "report")
    assert count >= 1
    assert plain.last_change.operation == "replace_text"
    assert (plain.last_change.text_before, plain.last_change.text_after) == (
        "document",
        "report",
    )
    assert len(plain.changes) == 1, "one report for the whole body"
    expected = plain.body.text

    package = sample("2010-sample1.docx")
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    package.change_tracking_mode = "TrackAll"
    assert package.body.replace_text("document", "report") == count
    assert package.body.text == expected, "the accepted view is the untracked result"
    assert len(package.body.get_tracked_changes()) == count * 2

    package.body.reject_all()
    assert package.body.text == sample("2010-sample1.docx").body.text


def test_replace_text_on_a_paragraph_and_on_a_range_are_tracked_too():
    _package, paragraph = tracked("one two one two one")
    assert paragraph.replace_text("one", "1") == 3
    assert paragraph.text == "1 two 1 two 1"

    _package2, paragraph2 = tracked("one two one two one")
    span = paragraph2.search("two one two")[0]
    assert span.replace_text("one", "1") == 1
    assert paragraph2.text == "one two 1 two one"
    assert [c.type for c in paragraph2.get_tracked_changes()] == ["Deleted", "Added"]


# ---------------------------------------------------------------------------
# what a tracked edit leaves untouched
# ---------------------------------------------------------------------------


def test_only_the_body_and_the_settings_part_are_rewritten_by_a_tracked_edit():
    package = sample("2010-sample1.docx")
    before = part_bytes(WordprocessingMLPackage.load(SAMPLES / "2010-sample1.docx").save())

    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    package.change_tracking_mode = "TrackAll"
    package.body.paragraphs[0].insert_text("New. ", location="Start")
    after = part_bytes(package.save())

    changed = [name for name, data in before.items() if after.get(name) != data]
    assert changed == ["word/document.xml", "word/settings.xml"]


def test_every_form_survives_a_save_and_a_reload():
    package, paragraph = tracked("Alpha beta gamma.")
    paragraph.replace_text("beta", "BETA")
    paragraph.font.bold = True
    paragraph.alignment = "Centered"
    package.body.insert_paragraph("Added.")
    table = package.body.insert_table(1, 1, values=[["cell"]])
    table.delete_rows(0)

    before = [(c.kind, c.type, c.author, c.text) for c in package.body.get_tracked_changes()]
    back = reloaded(package)
    after = [(c.kind, c.type, c.author, c.text) for c in back.body.get_tracked_changes()]
    assert after == before

    back.body.accept_all()
    assert back.body.paragraphs[0].text == "Alpha BETA gamma."


# ---------------------------------------------------------------------------
# the invariant: a rejected document is the document that was there
# (CR-003 section 16.12)
# ---------------------------------------------------------------------------

#: The documents the invariant runs over: one Word 2010 file with no
#: ``w14:paraId``, one with tracked changes of its own, and one Word stamps with
#: ``w14:paraId`` on every paragraph, which is what an agent's addresses rest on.
INVARIANT_DOCUMENTS = ("2010-sample1.docx", "sample-docx.docx", "DrawingML_GraphicData_wps.docx")


def every_tracked_verb(package, word: str = "the") -> None:
    """Every mutation Phase F tracks, over one body, in one session."""
    body = package.body
    body.replace_text(word, word.upper())
    body.paragraphs[0].insert_text("Start. ", location="Start")
    body.paragraphs[0].insert_text(" End.", location="End")
    body.paragraphs[0].insert_paragraph("Inserted after the first.", location="After")
    body.insert_paragraph("Appended at the end.")
    body.insert_markdown("## Heading\n\npara one\n\npara two\n")
    body.paragraphs[2].delete()
    body.paragraphs[0].font.bold = True
    body.paragraphs[0].alignment = "Centered"
    body.insert_table(2, 2, values=[["a", "b"], ["c", "d"]])
    table = body.tables[0]
    table.add_rows(1, values=[["x", "y"]])
    if table.row_count > 2:
        table.delete_rows(0)
    body.paragraphs[0].insert_inline_picture(_sample_png(), width=40)


def tracked_package(name: str):
    """One of ``samples/``, with the identity, the date and the seed fixed."""
    package = sample(name)
    package.id_seed = 20260917
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    package.change_tracking_mode = "TrackAll"
    return package


@pytest.mark.parametrize("name", INVARIANT_DOCUMENTS)
def test_rejecting_every_tracked_verb_gives_the_document_back(name):
    """The invariant that would have caught every defect of 16.9, 16.10 and 16.11.

    Every verb this phase tracks, over a **loaded** document, then
    ``reject_all()``: the body is **canonically** what it was, every paragraph
    id is still on its paragraph, no revision markup is left and
    ``describe()`` and ``outline()`` agree with it.
    """
    import canon  # scripts/ is on the path, from conftest

    control = sample(name)
    # ``samples/sample-docx.docx`` carries revisions of its own, and
    # ``reject_all()`` rejects those too --- so the document to compare with is
    # the source with **its** revisions rejected, which is what the control is
    control.body.reject_all()
    before = control.body.get_xml().encode("utf-8")
    ids_before = [p.para_id for p in control.body.paragraphs]

    package = tracked_package(name)
    package.body.reject_all()          # the document's own, as the control did
    package.change_tracking_mode = "TrackAll"
    every_tracked_verb(package)
    assert package.body.get_tracked_changes(), "the verbs really did write revisions"

    assert package.body.reject_all()
    after = package.body.get_xml().encode("utf-8")

    report = canon.compare(before, after)
    assert report.identical, f"{report.differences}: {[d.detail for d in report.diffs[:4]]}"
    assert package.body.get_tracked_changes() == []
    assert [p.para_id for p in package.body.paragraphs] == ids_before
    assert package.body.text == control.body.text
    assert package.outline().stats.tracked_changes == 0
    assert package.describe().authors["revisions"] == control.describe().authors["revisions"]


@pytest.mark.parametrize("name", INVARIANT_DOCUMENTS)
def test_accepting_every_tracked_verb_gives_what_the_untracked_run_gives(name):
    """The mirror: accepting leaves the document the same calls make untracked.

    The **text** and the paragraph ids, not the markup: an accepted insertion
    stays a run of its own where the untracked edit puts the characters into
    the run that was there, which is what Word leaves too (CR-003 section
    16.12). While the changes are *pending* the two differ by the paragraphs
    whose mark is deleted, which Word shows as well until they are accepted.
    """
    plain = sample(name)
    plain.id_seed = 20260917
    plain.body.reject_all()            # the document's own revisions, if any
    every_tracked_verb(plain)

    package = tracked_package(name)
    package.body.reject_all()
    package.change_tracking_mode = "TrackAll"
    every_tracked_verb(package)

    assert package.body.accept_all()
    assert package.body.get_tracked_changes() == []
    assert package.body.text == plain.body.text

    accepted_ids = [p.para_id for p in package.body.paragraphs]
    plain_ids = [p.para_id for p in plain.body.paragraphs]
    assert len(accepted_ids) == len(plain_ids)
    # at most one id differs, and it is the **mark join's**: accepting a deleted
    # paragraph mark keeps the first paragraph's ``w14:paraId`` and takes the
    # second's properties, as docx4j's ``AcceptTrackedChanges`` does, where an
    # untracked delete removes the first paragraph and keeps the second's id
    assert sum(a != b for a, b in zip(accepted_ids, plain_ids, strict=True)) <= 1
    xml = package.body.get_xml()
    assert "<w:ins " not in xml and "<w:del " not in xml
    assert "rPrChange" not in xml and "pPrChange" not in xml


def test_a_paragraph_appended_where_no_paragraph_can_carry_its_mark():
    """The one residue of the final-mark rule, pinned rather than forgotten.

    With no paragraph before it --- a body whose last block is a table, or
    ``samples/invoice2013.docx``, whose last blocks are bookmark elements ---
    there is no earlier mark for the break to live on, so nothing is marked and
    rejecting the insertion empties the paragraph rather than removing it
    (CR-003 sections 16.10 and 16.12).
    """
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude")
    package.tracked_change_date = WHEN
    package.body.insert_table(1, 1, values=[["a cell"]])
    package.change_tracking_mode = "TrackAll"

    made = package.body.insert_paragraph("Appended after a table.")
    assert made.element.p_pr is None, "no mark: there is none to move it to"
    assert "<w:ins " in made.get_xml(), "the runs are still an insertion"

    package.body.reject_all()
    assert [p.text for p in package.body.paragraphs] == ["a cell", ""]
    assert package.body.get_tracked_changes() == []


def _sample_png() -> bytes:
    """The PNG out of ``samples/Images.docx``, for the picture test."""
    with zipfile.ZipFile(SAMPLES / "Images.docx") as archive:
        name = next(n for n in archive.namelist() if n.endswith(".png"))
        return archive.read(name)
