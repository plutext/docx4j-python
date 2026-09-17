"""``List`` and ``ListItem``. CR-003 section 3.10, Phase H.

The views over ``w:num`` and ``w:numPr``: what they read, what the five verbs
write, and Word's rules for each (section 16.7's ``w:pPrChange``, the abstract
definition copied before a level setter touches a shared one, ``w:numId`` 0 for
style-contributed numbering). Every mutation is followed by a save, a reload and
a read back, as CR-003 section 7 requires, and the parts nothing touched are
still byte for byte.
"""

from __future__ import annotations

import datetime
import re

import pytest
from conftest import fixture, part_bytes, reloaded, sample

from docx4j_py import WordprocessingMLPackage, create_package
from docx4j_py.model.content import Author, List, ListItem
from docx4j_py.model.content.errors import ContentError

#: The date the tracked test fixes, so that the markup is reproducible.
WHEN = datetime.datetime(2026, 9, 17, 12, 0, tzinfo=datetime.UTC)


def numbered(*texts: str):
    """A created document with one numbered list of `texts`, level 0."""
    package = create_package()
    package.id_seed = 20260917
    body = package.body
    first = body.insert_paragraph(texts[0] if texts else "One")
    the_list = first.start_new_list()
    for text in texts[1:]:
        body.insert_paragraph(text).attach_to_list(the_list.id)
    return package, the_list


# ---------------------------------------------------------------------------
# reading (Office JS Word.List, Word.ListItem)
# ---------------------------------------------------------------------------


def test_a_bodys_lists_are_one_per_w_numId_in_document_order():
    package = fixture("lists.docx")
    lists = package.body.lists

    assert [one.id for one in lists] == [1]
    assert isinstance(lists[0], List)
    assert len(lists[0].paragraphs) == 4
    assert lists[0].level_types[0] == "Number"
    assert len(lists[0].level_types) == 9


def test_a_paragraph_knows_whether_it_is_a_list_item():
    package = fixture("lists.docx")
    items = [p for p in package.body.paragraphs if p.is_list_item]
    plain = [p for p in package.body.paragraphs if not p.is_list_item]

    assert len(items) == 4
    assert all(isinstance(p.list_item, ListItem) for p in items)
    assert all(p.list is not None and p.list.id == 1 for p in items)
    assert all(p.list_item is None and p.list is None for p in plain)


def test_list_string_is_the_label_word_paints():
    package, _the_list = numbered("One", "Two", "Three")

    assert [p.list_item.list_string for p in package.body.paragraphs] == ["1.", "2.", "3."]
    assert [p.list_item.sibling_index for p in package.body.paragraphs] == [0, 1, 2]
    assert [p.list_item.level for p in package.body.paragraphs] == [0, 0, 0]


def test_a_bullet_list_reports_a_bullet_character():
    package = create_package()
    package.id_seed = 20260917
    item = package.body.insert_paragraph("A point")
    the_list = item.start_new_list(kind="Bullet")

    assert the_list.level_types[0] == "Bullet"
    assert item.list_item.list_string == "•", "Word's Symbol w:lvlText, as a bullet"
    assert the_list.get_level_string(0) == "", "the level text itself is Word's own glyph"


def test_a_level_whose_format_is_none_has_no_label():
    package, the_list = numbered("One", "Two")
    the_list.set_level_numbering(0, "None")

    assert [p.list_item.list_string for p in package.body.paragraphs] == ["", ""]
    assert package.body.paragraphs[0].is_list_item, "it is still an item; it just has no label"


def test_the_levels_of_a_list_are_readable_one_by_one():
    package, the_list = numbered("One")
    package.body.insert_paragraph("Sub").attach_to_list(the_list.id, 1)

    assert the_list.level_exists(0) and the_list.level_exists(8)
    assert the_list.get_level_string(0) == "%1."
    assert [p.text for p in the_list.get_level_paragraphs(1)] == ["Sub"]
    assert [p.text for p in the_list.get_level_paragraphs(0)] == ["One"]


def test_an_ancestor_and_the_descendants_come_from_the_levels():
    package, the_list = numbered("One", "Two")
    body = package.body
    body.insert_paragraph("Two.a").attach_to_list(the_list.id, 1)
    body.insert_paragraph("Two.a.i").attach_to_list(the_list.id, 2)
    body.insert_paragraph("Three").attach_to_list(the_list.id, 0)

    two = body.paragraphs[1]
    deep = body.paragraphs[3]
    assert [p.text for p in two.list_item.get_descendants()] == ["Two.a", "Two.a.i"]
    assert [p.text for p in two.list_item.get_descendants(direct_children_only=True)] == ["Two.a"]
    assert deep.list_item.get_ancestor().text == "Two.a"
    assert deep.list_item.get_ancestor(parent_only=True).text == "Two.a"
    assert two.list_item.get_ancestor() is None, "a level-0 item has no ancestor"
    # the parent of a level-2 item under a level-0 one is not its parent
    assert body.paragraphs[4].list_item.get_descendants() == []


def test_to_dict_carries_the_list_item():
    package, the_list = numbered("One", "Two")
    paragraph = package.body.paragraphs[1]

    assert paragraph.to_dict()["list_item"] == {
        "level": 0,
        "list_string": "2.",
        "sibling_index": 1,
    }
    assert the_list.to_dict()["id"] == the_list.id
    assert the_list.to_dict()["level_types"][0] == "Number"
    assert the_list.to_dict()["address"] == f"list:{the_list.id}"
    assert package.body.paragraphs[0].to_dict().get("list_item") is not None


def test_a_paragraph_that_is_not_a_list_item_says_nothing_about_lists():
    package = create_package()
    paragraph = package.body.insert_paragraph("Plain")

    assert not paragraph.is_list_item
    assert paragraph.list is None and paragraph.list_item is None
    assert "list_item" not in paragraph.to_dict()
    assert package.body.lists == []


# ---------------------------------------------------------------------------
# start_new_list
# ---------------------------------------------------------------------------


def test_starting_a_list_creates_the_numbering_part_and_the_definitions():
    package = create_package()
    package.id_seed = 20260917
    assert package.numbering_definitions_part is None

    paragraph = package.body.insert_paragraph("One")
    the_list = paragraph.start_new_list()

    part = package.numbering_definitions_part
    assert part is not None
    numbering = part.contents
    assert len(numbering.abstract_num) == 1
    assert len(numbering.num) == 1
    assert int(numbering.num[0].num_id) == the_list.id
    assert the_list.id == 1, "the next free w:numId, above the document's own"
    assert paragraph.style_id == "ListParagraph", "as Word's ribbon applies it"

    back = reloaded(package)
    assert back.body.paragraphs[0].list_item.list_string == "1."
    assert back.body.lists[0].id == the_list.id


def test_starting_a_list_records_a_change_report():
    package = create_package()
    package.id_seed = 20260917
    paragraph = package.body.insert_paragraph("One")
    the_list = paragraph.start_new_list()

    report = package.last_change
    assert report.operation == "start_new_list"
    assert f"list:{the_list.id}" in report.addresses
    assert "/word/numbering.xml" in report.parts_touched
    assert "/word/document.xml" in report.parts_touched
    assert "/word/styles.xml" in report.parts_touched, "ListParagraph had to be defined"


def test_starting_a_list_on_an_item_is_refused_as_office_js_refuses_it():
    package, _the_list = numbered("One")

    with pytest.raises(ContentError) as raised:
        package.body.paragraphs[0].start_new_list()
    assert raised.value.code == "list.already_a_list_item"
    assert "detach_from_list" in raised.value.hint


def test_a_bullet_list_copies_docx4js_bullet_definition():
    package = create_package()
    package.id_seed = 20260917
    package.body.insert_paragraph("A point").start_new_list(kind="Bullet")

    levels = package.numbering_definitions_part.contents.abstract_num[0].lvl
    assert levels[0].num_fmt.val.value == "bullet"
    assert levels[0].r_pr.r_fonts.h_ansi == "Symbol"
    assert len(levels) == 9


def test_like_restarts_the_numbering_over_the_same_definition():
    package, first = numbered("One", "Two", "Three")
    body = package.body
    restarted = body.insert_paragraph("One again").start_new_list(like=first)

    numbering = package.numbering_definitions_part.contents
    assert len(numbering.abstract_num) == 1, "one definition, two instances"
    assert [int(n.num_id) for n in numbering.num] == [first.id, restarted.id]
    override = numbering.num[1].lvl_override[0]
    assert int(override.ilvl) == 0
    assert int(override.start_override.val) == 1

    back = reloaded(package)
    assert [p.list_item.list_string for p in back.body.paragraphs] == ["1.", "2.", "3.", "1."]


def test_a_paragraph_with_a_style_of_its_own_keeps_it():
    package = create_package()
    package.id_seed = 20260917
    heading = package.body.insert_paragraph("A numbered heading", style="Heading 1")
    heading.start_new_list()

    assert heading.style_id == "Heading1"
    assert heading.is_list_item


# ---------------------------------------------------------------------------
# attach, detach, level
# ---------------------------------------------------------------------------


def test_attaching_to_a_list_that_is_not_there_names_the_ones_that_are():
    package, the_list = numbered("One")

    with pytest.raises(ContentError) as raised:
        package.body.insert_paragraph("Two").attach_to_list(99)
    assert raised.value.code == "list.not_found"
    assert str(the_list.id) in raised.value.hint


def test_attaching_at_a_level_writes_the_ilvl():
    package, the_list = numbered("One")
    paragraph = package.body.insert_paragraph("Sub")
    paragraph.attach_to_list(the_list.id, 1)

    num_pr = paragraph.element.p_pr.num_pr
    assert int(num_pr.num_id.val) == the_list.id
    assert int(num_pr.ilvl.val) == 1
    assert package.last_change.operation == "attach_to_list"
    assert package.last_change.parts_touched == ("/word/document.xml",), (
        "nothing was written to the numbering part"
    )

    back = reloaded(package)
    assert back.body.paragraphs[1].list_item.level == 1


def test_the_level_setter_goes_through_the_paragraph_properties():
    package, _the_list = numbered("One", "Two")
    second = package.body.paragraphs[1]
    second.list_item.level = 1

    assert int(second.element.p_pr.num_pr.ilvl.val) == 1
    assert reloaded(package).body.paragraphs[1].list_item.list_string == "a."


def test_detaching_a_direct_numPr_removes_it():
    package, _the_list = numbered("One", "Two")
    second = package.body.paragraphs[1]
    second.detach_from_list()

    assert second.element.p_pr.num_pr is None
    assert not second.is_list_item
    back = reloaded(package)
    assert [p.is_list_item for p in back.body.paragraphs] == [True, False]
    assert back.body.paragraphs[0].list_item.list_string == "1."


def test_detaching_style_numbering_writes_w_numId_zero():
    """What Word writes to turn a numbered style's numbering off."""
    package, the_list = numbered("One")
    styles = package.style_definitions_part.contents
    numbered_style = next(s for s in styles.style if s.style_id == "ListParagraph")
    from docx4j_py.wml import el

    numbered_style.p_pr = numbered_style.p_pr or el.pPr()
    numbered_style.p_pr.num_pr = el.numPr(num_id=el.numId(val=the_list.id))
    from docx4j_py.model.listnumbering import invalidate

    invalidate(package)

    second = package.body.insert_paragraph("Two", style="ListParagraph")
    second.element.p_pr.num_pr = None  # numbered only through the style
    invalidate(package)
    assert second.is_list_item

    second.detach_from_list()
    assert int(second.element.p_pr.num_pr.num_id.val) == 0
    assert not second.is_list_item


def test_detaching_a_paragraph_that_is_not_an_item_does_nothing():
    package = create_package()
    paragraph = package.body.insert_paragraph("Plain")
    paragraph.detach_from_list()

    assert paragraph.element.p_pr is None or paragraph.element.p_pr.num_pr is None


# ---------------------------------------------------------------------------
# the level setters
# ---------------------------------------------------------------------------


def test_set_level_numbering_writes_the_numFmt_and_the_level_text():
    package, the_list = numbered("One", "Two", "Three")
    the_list.set_level_numbering(0, "LowerLetter", "%1)")

    assert [p.list_item.list_string for p in package.body.paragraphs] == ["a)", "b)", "c)"]
    assert the_list.get_level_string(0) == "%1)"
    assert package.last_change.operation == "list.set_level_numbering"
    assert "/word/numbering.xml" in package.last_change.parts_touched

    back = reloaded(package)
    assert [p.list_item.list_string for p in back.body.paragraphs] == ["a)", "b)", "c)"]


def test_set_level_numbering_refuses_a_value_office_js_does_not_have():
    _package, the_list = numbered("One")

    with pytest.raises(ContentError) as raised:
        the_list.set_level_numbering(0, "Klingon")
    assert raised.value.code == "list.numbering_invalid"


def test_a_level_setter_copies_a_definition_two_lists_share():
    package, first = numbered("One", "Two")
    second = package.body.insert_paragraph("Other").start_new_list(like=first)
    assert len(package.numbering_definitions_part.contents.abstract_num) == 1

    second.set_level_numbering(0, "UpperRoman")

    numbering = package.numbering_definitions_part.contents
    assert len(numbering.abstract_num) == 2, "the shared definition was copied first"
    assert [p.list_item.list_string for p in first.paragraphs] == ["1.", "2."]
    assert [p.list_item.list_string for p in second.paragraphs] == ["I."]


def test_set_level_bullet_writes_the_glyph_and_the_font():
    package, the_list = numbered("One")
    the_list.set_level_bullet(0, "Checkmark")

    lvl = package.numbering_definitions_part.contents.abstract_num[0].lvl[0]
    assert lvl.lvl_text.val == ""
    assert lvl.r_pr.r_fonts.h_ansi == "Wingdings"
    assert str(lvl.num_fmt.val) == "bullet"
    assert package.body.paragraphs[0].list_item.list_string == "✓"
    assert reloaded(package).body.lists[0].level_types[0] == "Bullet"


def test_a_custom_bullet_takes_a_char_code_and_a_font():
    _package, the_list = numbered("One")
    the_list.set_level_bullet(0, "Custom", char_code=0x2605, font_name="Arial")

    assert the_list.get_level_string(0) == "★"
    with pytest.raises(ContentError) as raised:
        the_list.set_level_bullet(0, "Custom")
    assert raised.value.code == "list.bullet_invalid"


def test_set_level_indents_is_in_points():
    package, the_list = numbered("One")
    the_list.set_level_indents(0, 36, 18)

    ind = package.numbering_definitions_part.contents.abstract_num[0].lvl[0].p_pr.ind
    assert int(ind.left) == 720
    assert int(ind.hanging) == 360
    assert reloaded(package).body.lists[0].definition.level(0).ind.left == 720


def test_a_level_the_definition_lacks_is_created_by_a_setter():
    package, the_list = numbered("One")
    numbering = package.numbering_definitions_part.contents
    del numbering.abstract_num[0].lvl[5:]
    from docx4j_py.model.listnumbering import invalidate

    invalidate(package)
    assert not the_list.level_exists(6)

    the_list.set_level_numbering(6, "Arabic")
    assert the_list.level_exists(6)
    assert [int(lvl.ilvl) for lvl in numbering.abstract_num[0].lvl] == [0, 1, 2, 3, 4, 6]


def test_a_level_outside_the_nine_is_refused():
    _package, the_list = numbered("One")

    with pytest.raises(ContentError) as raised:
        the_list.set_level_numbering(9, "Arabic")
    assert raised.value.code == "list.level_invalid"


# ---------------------------------------------------------------------------
# insert_paragraph
# ---------------------------------------------------------------------------


def test_inserting_at_the_start_and_the_end_of_a_list():
    package, the_list = numbered("One", "Two")
    the_list.insert_paragraph("Zero", "Start")
    the_list.insert_paragraph("Three", "End")

    assert [p.text for p in package.body.paragraphs] == ["Zero", "One", "Two", "Three"]
    assert [p.list_item.list_string for p in package.body.paragraphs] == [
        "1.",
        "2.",
        "3.",
        "4.",
    ]
    back = reloaded(package)
    assert [p.text for p in back.body.lists[0].paragraphs] == ["Zero", "One", "Two", "Three"]


def test_a_new_item_is_at_level_zero_whatever_the_item_beside_it_is():
    package, the_list = numbered("One")
    package.body.insert_paragraph("Sub").attach_to_list(the_list.id, 1)
    the_list.insert_paragraph("Two", "End")

    assert [p.list_item.level for p in package.body.paragraphs] == [0, 1, 0]


def test_inserting_into_a_list_with_no_items_here_is_refused():
    package, the_list = numbered("One")
    package.body.paragraphs[0].detach_from_list()

    with pytest.raises(ContentError) as raised:
        the_list.insert_paragraph("x", "End")
    assert raised.value.code == "list.empty"


# ---------------------------------------------------------------------------
# tracking (CR-003 section 16.7)
# ---------------------------------------------------------------------------


def test_attaching_while_tracking_records_a_pPrChange_and_a_reject_undoes_it():
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude", initials="C")
    package.tracked_change_date = WHEN
    body = package.body
    first = body.insert_paragraph("One")
    the_list = first.start_new_list()
    plain = body.insert_paragraph("Not an item yet")
    before = plain.get_xml()

    package.change_tracking_mode = "TrackAll"
    plain.attach_to_list(the_list.id)

    assert plain.element.p_pr.p_pr_change is not None, "Word writes w:pPrChange for this"
    assert "w:numberingChange" not in plain.get_xml(), "the deprecated revision is never written"
    changes = package.body.get_tracked_changes()
    assert [c.type for c in changes] == ["Formatted"]

    back = reloaded(package)
    assert back.body.reject_all() == 1
    rejected = back.body.paragraphs[1]
    assert not rejected.is_list_item
    assert "w:pPrChange" not in rejected.get_xml()
    assert "w:numPr" not in rejected.get_xml(), "the markup is gone, not just the value"
    assert _canonical(rejected.get_xml()) == _canonical(before)


def test_starting_a_list_on_a_paragraph_this_author_inserted_records_nothing():
    package = create_package()
    package.id_seed = 20260917
    package.author = Author("Claude", initials="C")
    package.tracked_change_date = WHEN
    package.change_tracking_mode = "TrackAll"

    paragraph = package.body.insert_paragraph("One")
    paragraph.start_new_list()

    assert paragraph.element.p_pr.p_pr_change is None, "its properties are this author's too"
    assert [c.type for c in package.body.get_tracked_changes()] == ["Added"]


def test_nothing_in_the_numbering_part_is_tracked():
    package, the_list = numbered("One")
    package.author = Author("Claude", initials="C")
    package.change_tracking_mode = "TrackAll"
    the_list.set_level_numbering(0, "UpperRoman")

    numbering = package.numbering_definitions_part.get_xml()
    assert "w:ins" not in numbering and "w:del" not in numbering


def _canonical(xml: str) -> str:
    """The paragraph's markup, with the ids and the whitespace out of the way."""
    return re.sub(r"\s+", " ", re.sub(r'w14:paraId="[^"]*"', "", xml)).strip()


# ---------------------------------------------------------------------------
# what is touched, round trips and dry runs
# ---------------------------------------------------------------------------


def test_starting_a_list_in_a_loaded_document_leaves_the_other_parts_alone():
    package = sample("2010-sample1.docx")
    package.id_seed = 20260917
    _ = package.body, package.style_definitions_part.contents
    before = part_bytes(package.save())

    package.body.paragraphs[0].start_new_list()
    after = part_bytes(package.save())

    changed = sorted(name for name, data in before.items() if after.get(name) != data)
    assert changed == ["[Content_Types].xml", "word/document.xml", "word/styles.xml"], (
        "the content types gained the numbering part's override"
    )
    assert "word/numbering.xml" in after, "and the numbering part was created"


def test_a_dry_run_of_starting_a_list_leaves_the_document_byte_for_byte():
    package = sample("2010-sample1.docx")
    package.id_seed = 20260917
    _ = package.body, package.style_definitions_part.contents
    before = part_bytes(package.save(), relationships=True)

    with package.dry_run() as trial:
        trial.body.paragraphs[0].start_new_list()
        assert trial.body.paragraphs[0].is_list_item
        preview = trial.last_change.to_dict()

    after = part_bytes(package.save(), relationships=True)
    assert [n for n, d in before.items() if after.get(n) != d] == []
    assert sorted(after) == sorted(before), "the numbering part the trial made is gone"
    assert not package.body.paragraphs[0].is_list_item
    assert "/word/numbering.xml" in preview["parts_touched"]


def test_a_two_level_list_with_a_restart_round_trips_through_markdown():
    package, first = numbered("Alpha", "Beta")
    body = package.body
    body.insert_paragraph("Beta.a").attach_to_list(first.id, 1)
    body.insert_paragraph("Restarted").start_new_list(like=first)

    markdown = body.to_markdown()
    assert markdown == "1. Alpha\n2. Beta\n   1. Beta.a\n\n1) Restarted"

    again = create_package()
    again.id_seed = 20260917
    again.body.insert_markdown(markdown + "\n")
    assert [p.text for p in again.body.paragraphs] == ["Alpha", "Beta", "Beta.a", "Restarted"]
    assert [p.list_item.level for p in again.body.paragraphs] == [0, 0, 1, 0]
    assert [p.list_item.list_string for p in again.body.paragraphs] == ["1.", "2.", "1.", "1."]
    assert again.body.to_markdown() == markdown


def test_a_loose_items_follow_on_paragraph_is_rendered_under_it():
    """CR-003 section 13.4, decided in Phase H: Java's rule, implemented."""
    package = create_package()
    package.id_seed = 20260917
    package.body.insert_markdown("- one\n\n  continued here\n- two\n")

    assert package.body.to_markdown() == "- one\n\n  continued here\n- two"


def test_the_same_seed_and_the_same_calls_give_the_same_bytes():
    def build() -> dict[str, bytes]:
        package, the_list = numbered("One", "Two")
        the_list.set_level_numbering(0, "LowerLetter", "%1)")
        package.body.insert_paragraph("Three").attach_to_list(the_list.id, 1)
        return part_bytes(package.save())

    first, second = build(), build()
    for name, data in first.items():
        if name.startswith("docProps"):
            continue  # dcterms:created is the time of day (CR-002 section 12.8)
        assert second[name] == data, name


def test_a_list_survives_a_save_and_a_reload_in_a_loaded_document():
    package = sample("2010-sample1.docx")
    package.id_seed = 20260917
    first = package.body.paragraphs[0]
    the_list = first.start_new_list()
    package.body.paragraphs[1].attach_to_list(the_list.id, 1)

    back = WordprocessingMLPackage.load(package.save())
    items = [p for p in back.body.paragraphs if p.is_list_item]
    assert [p.list_item.list_string for p in items] == ["1.", "a."]
    assert back.body.lists[0].id == the_list.id
