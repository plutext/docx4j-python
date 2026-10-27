"""Word's table rule: a cell ends with a paragraph, as Word reads it (CR-016 Phase A).

    .venv-fork/bin/python -m pytest tests/content/test_table_rules.py

docx4j's ``TableRulesTest`` (``18a3a31f4``), its four tests and their probes,
then docx4j-core-ts's ``test/table-rules.test.mjs`` at ``222cc20``, its seven
tests, their strings kept. core-ts's first two are its port of docx4j's four,
and are those here. The other five hold the edits of core-ts CR-002 section
55.2's table to section 55.6's answers, which are Word 365's in check 52's
cases 86 to 95 (``test_ranges.py`` holds the port's run of those cases to
Word's run, its save and its Accept All). Where this package's answer differs
from core-ts's the test asserts this package's and says why (CR-016 section
4.9). Then this package's own: a tracked ``TableCell.value`` (Q11), the
controls', the kept branches' and a broken cell's cases, the python-docx
facade's ``_Cell.add_table`` (Q13), and one ``ChangeReport`` per write with
its ``dry_run`` (CR-003 sections 3.4 and 12.5).
"""

from __future__ import annotations

import re
from typing import Any

import pytest

from docx4j_py import create_package, load
from docx4j_py.model import revisions
from docx4j_py.model.content import (
    AddressError,
    end_with_paragraph,
    ends_with_paragraph,
)
from docx4j_py.traversal import element_name
from docx4j_py.wml import Tc, to_xml, wml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# --- docx4j's TableRulesTest -------------------------------------------------------------------------------------

NS = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w99x="urn:docx4j:probe:unknown"'
P = "<w:p><w:r><w:t>p</w:t></w:r></w:p>"
TBL = f'<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>{P}</w:tc></w:tr></w:tbl>'


def sdt(content: str) -> str:
    return f'<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>'


def mc(requires: str, choice: str, fallback: str) -> str:
    return (
        f'<mc:AlternateContent><mc:Choice Requires="{requires}">{choice}</mc:Choice>'
        f"<mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"
    )


def tc(content: str) -> Tc:
    """A ``w:tc`` holding `content`, unmarshalled as it is (docx4j's ``XmlUtils.unmarshalString``): nothing repaired."""
    cell = wml(f'<w:tc {NS}><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>{content}</w:tc>')
    assert isinstance(cell, Tc)
    return cell


def opens(content: str) -> None:
    assert ends_with_paragraph(tc(content)), content
    assert ends_with_paragraph(tc(content).content), "the cell or its list, docx4j's two overloads"


def unreadable(content: str) -> None:
    assert not ends_with_paragraph(tc(content)), content
    assert not ends_with_paragraph(tc(content).content), content


def test_the_probes_word_opens():
    opens(TBL + P)  # 01
    opens(TBL + sdt(P))  # 03
    opens(TBL + sdt(sdt(P)))  # 04
    opens(TBL + mc("w14", P, P))  # 06
    opens(TBL + mc("w99x", P, P))  # 07, the Fallback read
    opens(TBL + mc("w14", P, TBL))  # 08, the unread Fallback not checked
    opens(sdt(P))  # 10
    opens(mc("w14", P, P))  # 11


def test_the_probes_word_repairs():
    unreadable(TBL)  # 02
    unreadable(TBL + sdt(TBL))  # 05
    unreadable(TBL + mc("w14", TBL, P))  # 09, the read Choice ends in a table
    unreadable(sdt(TBL))  # 12
    unreadable("")  # no block at all (rule 1)


def test_range_markup_after_the_paragraph_does_not_count():
    opens(f'{P}<w:bookmarkEnd w:id="0"/>')
    unreadable(f'{TBL}<w:bookmarkEnd w:id="0"/>')


def test_end_with_paragraph_appends_once_after_the_wrapper():
    cell = tc(TBL + sdt(TBL))
    appended = end_with_paragraph(cell)
    assert appended is not None, "docx4j's true"
    xml = to_xml(cell)
    assert xml.endswith("</w:sdt><w:p/></w:tc>"), xml
    assert end_with_paragraph(cell) is None, "already ends with a paragraph"
    assert len(cell.content) == 3
    # core-ts's: the paragraph appended is the cell's own child, its parent linked
    assert cell.content[2] is appended and appended.parent is cell

    wrapped = tc(TBL + sdt(P))
    assert end_with_paragraph(wrapped) is None
    assert len(wrapped.content) == 2


# --- core-ts's table-rules.test.mjs: the edits ---------------------------------------------------------------------

#: The cell's paragraph: a heading, centred, its text underlined (check 52's was a list item, underlined).
CELL_P = (
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>'
    '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>CELL</w:t></w:r></w:p>'
)
NESTED = (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>NESTED</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
)
NEW_TABLE = NESTED.replace("NESTED", "NEWTABLE")


def cell_holding(content: str):
    """A new document whose one table's first cell holds `content` (p: the paragraph; tp: a nested table, then it)."""
    package = create_package()
    package.body.insert_xml(
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr></w:tbl>',
        location="End",
    )
    return package, lambda: package.body.tables[0].cell(0, 0)


def _local(item: Any) -> str:
    return (element_name(item) or "").rsplit("}", 1)[-1]


def _val(value: Any) -> Any:
    found = getattr(value, "val", None)
    return getattr(found, "value", found)


def runs(items: Any) -> str:
    """A run list's text, a revision holder's as ``ins:`` or ``del:``."""
    out = []
    for item in items or ():
        name = _local(item)
        if name in ("ins", "del"):
            out.append(f"{name}:{runs(item.content)}")
        elif name == "r":
            out.append("".join(getattr(c, "value", "") or "" for c in item.content or () if _local(c) in ("t", "delText")))
    return "".join(out)


def shape(items: Any) -> list[str]:
    """A cell's blocks: a paragraph as its style, alignment and mark (``u``, ``ins``, ``del``), then its text; a table; a control."""
    out = []
    for item in items or ():
        name = _local(item)
        if name == "p":
            p_pr = item.p_pr
            r_pr = getattr(p_pr, "r_pr", None)
            marks = (("u", "u"), ("ins", "ins"), ("del", "del_value"))
            mark = ",".join(key for key, field in marks if getattr(r_pr, field, None) is not None)
            props = ",".join(str(v) for v in (_val(getattr(p_pr, "p_style", None)), _val(getattr(p_pr, "jc", None))) if v)
            out.append(f"P[{props}|{mark}] {runs(item.content)}".rstrip())
        elif name == "tbl":
            rows = [row for row in item.content or () if _local(row) == "tr"]
            ins = "+ins" if any(getattr(getattr(r, "tr_pr", None), "ins", None) is not None for r in rows) else ""
            deleted = "+del" if any(getattr(getattr(r, "tr_pr", None), "del_value", None) is not None for r in rows) else ""
            out.append(f"TBL{ins}{deleted}")
        elif name == "sdt":
            out.append("SDT{" + " ".join(shape(item.sdt_content.content)) + "}")
        else:
            out.append(f"<{name}>")
    return out


def cell_shape(cell) -> list[str]:
    return shape(cell().body.content)


def first_cell(package):
    return package.body.tables[0].cell(0, 0)


def resolved(package, verb: str):
    """The package saved, loaded again and every change accepted or rejected, as core-ts's tests resolve a save."""
    again = load(package.save())
    (revisions.accept_all if verb == "accept" else revisions.reject_all)(again)
    return again


def test_an_edit_that_would_take_the_paragraph_a_cell_ends_with_keeps_it_emptied_its_properties_kept_check_52_cases_88_89_92_94_and_95():
    for content, before in ((CELL_P, []), (NESTED + CELL_P, ["TBL"])):
        _package, cell = cell_holding(content)
        cell().body.paragraphs[-1].delete()
        assert cell_shape(cell) == [*before, "P[Heading1,center|]"], "Paragraph.delete"

        _package, cell = cell_holding(content)
        cell().body.clear()
        assert cell_shape(cell) == ["P[Heading1,center|]"], "Body.clear"

        _package, cell = cell_holding(content)
        cell().body.insert_ooxml(NEW_TABLE, location="Replace")
        assert cell_shape(cell) == ["TBL", "P[Heading1,center|]"], "Body.insert_ooxml(table, 'Replace')"

        _package, cell = cell_holding(content)
        cell().body.paragraphs[-1].insert_ooxml(NEW_TABLE, location="Replace")
        assert cell_shape(cell) == [*before, "TBL", "P[Heading1,center|]"], "Paragraph.insert_ooxml(table, 'Replace')"
    # a control around the cell's paragraph, deleted with its content: the paragraph stays, outside it (case 95)
    _package, cell = cell_holding(CELL_P)
    cell().body.insert_content_control()
    assert cell_shape(cell) == ["SDT{P[Heading1,center|] CELL}"]
    cell().body.content_controls[0].delete(keep_content=False)
    assert cell_shape(cell) == ["P[Heading1,center|]"]
    _package, cell = cell_holding(NESTED + CELL_P)
    cell().body.paragraphs[-1].insert_content_control()
    cell().body.content_controls[0].delete(keep_content=False)
    assert cell_shape(cell) == ["TBL", "P[Heading1,center|]"]
    # where the paragraph before keeps the rule, the last one goes, as before
    _package, cell = cell_holding(CELL_P.replace("CELL", "FIRST") + CELL_P)
    cell().body.paragraphs[-1].delete()
    assert cell_shape(cell) == ["P[Heading1,center|] FIRST"]
    # a Replace whose content ends with a paragraph replaces the cell's paragraph, as before
    _package, cell = cell_holding(CELL_P)
    cell().body.insert_ooxml(f'<w:p xmlns:w="{W}"><w:r><w:t>NEW</w:t></w:r></w:p>', location="Replace")
    assert cell_shape(cell) == ["P[|] NEW"]


def test_tracked_a_cells_last_paragraph_has_its_text_deleted_and_its_mark_kept_accept_all_leaves_it_emptied_reject_all_as_it_was_check_52_cases_90_91_and_93():
    for content, before in ((CELL_P, []), (NESTED + CELL_P, ["TBL"])):
        for edit in ("delete", "clear"):
            package, cell = cell_holding(content)
            package.change_tracking_mode = "TrackAll"
            if edit == "delete":
                cell().body.paragraphs[-1].delete()
            else:
                cell().body.clear()
            package.change_tracking_mode = "Off"
            # a tracked clear deletes the nested table's rows too
            nested = ["TBL+del" for _ in before] if edit == "clear" else before
            label = f"{edit} {'p' if content == CELL_P else 'tp'}"
            assert cell_shape(cell) == [*nested, "P[Heading1,center|] del:CELL"], label
            accepted = [] if edit == "clear" else before
            assert shape(first_cell(resolved(package, "accept")).body.content) == [*accepted, "P[Heading1,center|]"], f"{label}, accepted"
            assert shape(first_cell(resolved(package, "reject")).body.content) == [*before, "P[Heading1,center|] CELL"], f"{label}, rejected"


INSERTIONS = {
    "insert_table": lambda cell: cell().body.insert_table(1, 1, location="End", values=[["NEWTABLE"]]),
    "insert_xml": lambda cell: cell().body.insert_xml(NEW_TABLE, location="End"),
    "insert_ooxml After": lambda cell: cell().body.paragraphs[-1].insert_ooxml(NEW_TABLE, location="After"),
    # this package's own verbs that put blocks at a cell's end the same way (CR-016 section 2.1)
    "insert_markdown": lambda cell: cell().body.insert_markdown("| NEWTABLE |\n|---|\n| ROW |\n", location="End"),
    "Paragraph.insert_markdown After": lambda cell: cell().body.paragraphs[-1].insert_markdown("| NEWTABLE |\n|---|\n| ROW |\n"),
    "insert_element": lambda cell: cell().body.insert_element(wml(NEW_TABLE), location="End"),
    "insert_xml After": lambda cell: cell().body.paragraphs[-1].insert_xml(NEW_TABLE, location="After"),
    "Range.insert_ooxml After": lambda cell: cell().body.paragraphs[-1].get_range("Whole").insert_ooxml(NEW_TABLE, location="After"),
}


@pytest.mark.parametrize("name", list(INSERTIONS))
def test_a_table_at_a_cells_end_the_cells_paragraph_is_split_at_its_end_and_its_mark_stays_last_after_the_table_check_52_cases_86_and_87(name):
    insert = INSERTIONS[name]
    package, cell = cell_holding(CELL_P)
    insert(cell)
    # the first half's mark in its text's formatting (case 86's w:u), the cell's mark with its properties after the table
    assert cell_shape(cell) == ["P[Heading1,center|u] CELL", "TBL", "P[Heading1,center|]"], name
    package, cell = cell_holding(CELL_P)
    package.change_tracking_mode = "TrackAll"
    insert(cell)
    package.change_tracking_mode = "Off"
    # tracked, the inserted mark is the first half's, with the rows; the cell's own mark is not marked (case 87)
    assert cell_shape(cell) == ["P[Heading1,center|u,ins] CELL", "TBL+ins", "P[Heading1,center|]"], f"{name}, tracked"
    accepted = shape(first_cell(resolved(package, "accept")).body.content)
    assert accepted == ["P[Heading1,center|u] CELL", "TBL", "P[Heading1,center|]"], f"{name}, accepted"
    assert shape(first_cell(resolved(package, "reject")).body.content) == ["P[Heading1,center|] CELL"], f"{name}, rejected"


def test_at_the_start_the_rule_holds_and_nothing_is_added_a_range_replaced_by_a_table_keeps_this_packages_form():
    _package, cell = cell_holding(CELL_P)
    cell().body.insert_table(1, 1, location="Start")
    assert cell_shape(cell) == ["TBL", "P[Heading1,center|] CELL"]
    # core-ts: "the paragraph keeps its text, so a range over it still holds it: Range.insertOoxml(table,
    # 'Replace'), unmeasured in Word, puts the table after it, then takes its text", giving
    # ['P[Heading1,center|u]', 'TBL', 'P[Heading1,center|]']. This package's Range.insert_ooxml takes the text
    # first and puts the blocks before the paragraph it leaves (CR-013 Phase C), which keeps the rule and is
    # Word's form for Body.insertOoxml (case 94): a recorded departure (CR-016 section 10)
    for where in ("Content", "Whole"):
        _package, cell = cell_holding(CELL_P)
        cell().body.paragraphs[0].get_range(where).insert_ooxml(NEW_TABLE, location="Replace")
        assert cell_shape(cell) == ["TBL", "P[Heading1,center|]"], where


def test_elsewhere_docx4js_repair_a_table_at_a_controls_end_in_a_cell_the_cells_of_inserted_content_d3_what_kept_the_rule_still_keeps_it():
    # a table put at the end of a control's content, the control the cell's last block: an empty paragraph after the control
    package, cell = cell_holding(CELL_P)
    control = cell().body.insert_content_control()
    control.paragraphs[0].insert_ooxml(NEW_TABLE, location="After")
    assert cell_shape(cell) == ["SDT{P[Heading1,center|] CELL TBL}", "P[|]"]
    # D3: a table whose cell ends with a table, inserted, has that cell repaired; the outer cell keeps its paragraph
    package, cell = cell_holding(CELL_P)
    inner = (
        '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>'
        f"</w:tcPr>{NESTED}</w:tc></w:tr></w:tbl>"
    )
    package.body.insert_xml(inner, location="End")
    assert shape(package.body.tables[1].cell(0, 0).body.content) == ["TBL", "P[|]"]
    # what kept the rule: a cell's value, Range.delete of the whole range, insert_text Replace, a tracked value
    package, cell = cell_holding(NESTED + CELL_P)
    cell().value = "VALUE"
    assert cell_shape(cell) == ["P[|] VALUE"]
    package, cell = cell_holding(CELL_P)
    cell().body.paragraphs[0].get_range("Whole").delete()
    assert cell_shape(cell) == ["P[Heading1,center|]"]
    package, cell = cell_holding(CELL_P)
    cell().body.insert_text("TEXT", location="Replace")
    assert cell_shape(cell) == ["P[|] TEXT"]
    package, cell = cell_holding(CELL_P)
    package.change_tracking_mode = "TrackAll"
    cell().value = "VALUE"
    package.change_tracking_mode = "Off"
    revisions.accept_all(package)
    assert cell_shape(cell) == ["P[|] VALUE"]


def test_d4_a_file_whose_cell_ends_with_a_table_is_left_as_it_is_until_an_edit_touches_that_cell():
    package, cell = cell_holding(CELL_P)
    # made broken behind the API's back, as a file from elsewhere is, and saved as it is: the save's check
    # (tests/conftest.py, CR-020 Phase B) would refuse the cell this session broke
    cell().body.content.append(wml(NESTED))
    again = load(package.save(check=False))
    assert shape(first_cell(again).body.content) == ["P[Heading1,center|] CELL", "TBL"]
    # an untouched part is written back byte for byte, broken as it was read
    assert load(again.save()).main_document_part.xml == again.main_document_part.xml
    # an edit in that cell repairs it, docx4j's way
    first_cell(again).body.insert_paragraph("FIRST", location="Start")
    assert shape(first_cell(again).body.content) == ["P[|] FIRST", "P[Heading1,center|] CELL", "TBL", "P[|]"]


# --- this package's own ------------------------------------------------------------------------------------------


@pytest.mark.parametrize("content", [CELL_P, NESTED + CELL_P], ids=["p", "tp"])
def test_q11_a_tracked_cell_value_deletes_the_old_content_as_a_tracked_change_and_reject_all_gives_it_back(content):
    package, cell = cell_holding(content)
    original = cell_shape(cell)
    package.change_tracking_mode = "TrackAll"
    cell().value = "VALUE"
    package.change_tracking_mode = "Off"
    # the old content a tracked deletion (its mark too, with the mark the append moved back onto it, CR-003 section
    # 16.10), then the new paragraph; 0.1.0 and 0.2.0 before CR-016 removed the old content untracked
    nested = ["TBL+del"] if content != CELL_P else []
    assert cell_shape(cell) == [*nested, "P[Heading1,center|ins,del] del:CELL", "P[|] ins:VALUE"]
    types = sorted({change.type for change in cell().body.get_tracked_changes()})
    assert "Deleted" in types and "Added" in types, types
    assert shape(first_cell(resolved(package, "accept")).body.content) == ["P[|] VALUE"]
    assert shape(first_cell(resolved(package, "reject")).body.content) == original


def test_a_controls_delete_keeping_its_content_leaves_the_paragraph_and_a_cell_left_ending_in_a_table_is_repaired():
    # keep_content=True (the default): the paragraph comes out of the control, as before
    _package, cell = cell_holding(CELL_P)
    cell().body.insert_content_control()
    cell().body.content_controls[0].delete()
    assert cell_shape(cell) == ["P[Heading1,center|] CELL"]
    # a control holding only a table, the cell's last block (broken as read): unwrapped, the cell is repaired
    _package, cell = cell_holding(CELL_P)
    cell().body.content.append(wml(sdt(NESTED), wrapper="body"))
    assert not ends_with_paragraph(cell().element)
    cell().body.content_controls[0].delete()
    assert cell_shape(cell) == ["P[Heading1,center|] CELL", "TBL", "P[|]"]
    # with its content, the cell's paragraph before it keeps the rule: nothing is added
    _package, cell = cell_holding(CELL_P)
    cell().body.content.append(wml(sdt(NESTED), wrapper="body"))
    cell().body.content_controls[0].delete(keep_content=False)
    assert cell_shape(cell) == ["P[Heading1,center|] CELL"]


def test_a_control_in_a_cell_whose_body_is_cleared_or_replaced_keeps_the_cells_paragraph_inside_it():
    package, cell = cell_holding(CELL_P)
    control = cell().body.insert_content_control()
    control.body().clear()
    assert cell_shape(cell) == ["SDT{P[Heading1,center|]}"]
    package, cell = cell_holding(CELL_P)
    control = cell().body.insert_content_control()
    control.body().insert_ooxml(NEW_TABLE, location="Replace")
    assert cell_shape(cell) == ["SDT{TBL P[Heading1,center|]}"]
    # outside a cell nothing is kept: the body's own rules (a control in the document's body)
    package = create_package()
    package.body.insert_paragraph("BODY")
    control = package.body.insert_content_control()
    control.body().clear()
    assert shape(package.body.content) == ["SDT{}"]


def test_a_kept_mc_alternate_content_is_read_on_words_branch_cleared_its_paragraph_is_kept_and_moved_out_an_edit_inside_it_refused():
    content = mc("w14", CELL_P, CELL_P.replace("CELL", "FALLBACK"))
    package, cell = cell_holding(content)
    assert ends_with_paragraph(cell().element)
    choice = cell().element.content[0].choice[0].any_element[0]
    # an edit inside the kept branch stays refused (CR-011 Phase F, CR-012 Phase A), as before
    with pytest.raises(AddressError) as refused:
        cell().body.paragraphs[0].delete()
    assert refused.value.code == "address.read_only"
    # a clear takes the element whole and keeps Word's branch's paragraph, emptied, moved out (core-ts's)
    cell().body.clear()
    assert cell_shape(cell) == ["P[Heading1,center|]"]
    assert cell().element.content[0] is choice and choice.parent is cell().element
    # tracked, a kept element is refused as before (kept_content_under)
    package, cell = cell_holding(content)
    package.change_tracking_mode = "TrackAll"
    with pytest.raises(AddressError) as refused:
        cell().body.clear()
    assert refused.value.code == "address.read_only"


def test_an_inserted_cell_that_breaks_the_rule_is_repaired_before_it_goes_in_and_under_tracking_the_repair_is_inserted_with_it():
    package = create_package()
    package.body.insert_paragraph("before")
    package.change_tracking_mode = "TrackAll"
    package.body.insert_xml(f"<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc>{NESTED}</w:tc></w:tr></w:tbl>")
    package.change_tracking_mode = "Off"
    cell = package.body.tables[0].cell(0, 0)
    # D3, before the insertion: the repaired cell's new paragraph is part of what was inserted
    assert shape(cell.body.content) == ["TBL+ins", "P[|ins]"]
    assert shape(first_cell(resolved(package, "accept")).body.content) == ["TBL", "P[|]"]
    assert resolved(package, "reject").body.tables == []


def test_a_repair_after_an_insert_is_plain_under_tracking():
    # a table put after a control that ends the cell: docx4j's repair, an empty paragraph after the control, unmarked
    package, cell = cell_holding(CELL_P)
    control = cell().body.insert_content_control()
    package.change_tracking_mode = "TrackAll"
    control.paragraphs[0].insert_ooxml(NEW_TABLE, location="After")
    package.change_tracking_mode = "Off"
    assert cell_shape(cell) == ["SDT{P[Heading1,center|] CELL TBL+ins}", "P[|]"]
    assert shape(first_cell(resolved(package, "reject")).body.content) == ["SDT{P[Heading1,center|] CELL}", "P[|]"]


def test_the_rule_reads_through_custom_xml_and_a_range_marker_and_not_into_a_nested_tables_cells():
    custom = f'<w:customXml w:element="item">{CELL_P}</w:customXml>'
    assert ends_with_paragraph(tc(NESTED + custom))
    assert not ends_with_paragraph(tc(f'<w:customXml w:element="item">{NESTED}</w:customXml>'))
    assert ends_with_paragraph(tc(f'{NESTED}{P}<w:commentRangeEnd w:id="1"/><w:bookmarkEnd w:id="0"/>'))
    # a nested table's own last paragraph is not the outer cell's end
    assert not ends_with_paragraph(tc(f"{P}{NESTED}"))


def test_the_facades_cell_add_table_is_python_docxs_the_cells_paragraph_untouched_the_table_then_one_plain_paragraph():
    from docx4j_py.docx import Document

    document = Document()
    outer = document.add_table(1, 1).cell(0, 0)
    outer.text = "CELL"
    inner = outer.add_table(2, 2)
    blocks = [_local(item) for item in outer._tc.content if _local(item) in ("p", "tbl")]
    assert blocks == ["p", "tbl", "p"], "python-docx 1.2.0: the paragraph, the table, one paragraph"
    first, last = outer.paragraphs
    assert first.text == "CELL" and first._p.p_pr is None, "the cell's own paragraph untouched"
    assert last.text == "" and last._p.p_pr is None and not last._p.content, "one plain <w:p/>"
    assert len(inner.rows) == 2 and len(inner.columns) == 2
    assert outer.tables[0]._tbl is inner._tbl
    # no split of the cell's paragraph stacked on it (CR-016 Q13), so nothing beyond python-docx's one paragraph
    again = load(document.save())
    cell = again.body.tables[0].cell(0, 0)
    assert [_local(item) for item in cell.body.content] == ["p", "tbl", "p"]


def _cell_writes() -> dict[str, tuple[str, str, Any]]:
    """Each write of Phase A in a cell: (its cell's content, the ChangeReport's operation, the write)."""
    return {
        "Paragraph.delete": (CELL_P, "delete", lambda pkg: first_cell(pkg).body.paragraphs[-1].delete()),
        "Body.clear": (NESTED + CELL_P, "clear", lambda pkg: first_cell(pkg).body.clear()),
        "insert_ooxml Replace": (CELL_P, "insert_ooxml", lambda pkg: first_cell(pkg).body.insert_ooxml(NEW_TABLE, location="Replace")),
        "insert_table End": (CELL_P, "insert_table", lambda pkg: first_cell(pkg).body.insert_table(1, 1, location="End")),
        "insert_markdown End": (CELL_P, "insert_markdown", lambda pkg: first_cell(pkg).body.insert_markdown("| A |\n|---|\n| B |\n")),
        "ContentControl.delete": (
            sdt(CELL_P),
            "delete",
            lambda pkg: first_cell(pkg).body.content_controls[0].delete(keep_content=False),
        ),
        "TableCell.value": (NESTED + CELL_P, "set_value", lambda pkg: setattr(first_cell(pkg), "value", "VALUE")),
    }


#: Each write untracked and tracked, but a control's delete, which is not tracked (CR-003 section 4).
_REPORTED = [
    pytest.param(write, tracked, id=f"{write}-{'tracked' if tracked else 'untracked'}")
    for write in _cell_writes()
    for tracked in (False, True)
    if not (tracked and write == "ContentControl.delete")
]


@pytest.mark.parametrize(("write", "tracked"), _REPORTED)
def test_each_write_records_one_change_report_and_a_dry_run_leaves_the_package_as_it_was(write, tracked):
    content, operation, edit = _cell_writes()[write]
    package, _cell = cell_holding(content)
    if tracked:
        package.change_tracking_mode = "TrackAll"
    before = package.main_document_part.xml
    with package.dry_run() as trial:
        edit(trial)
        assert len(trial.changes) == 1
        assert trial.changes[-1].operation == operation
        assert ends_with_paragraph(first_cell(trial).element)
    assert package.main_document_part.xml == before, "the real package is what it was"
    count = len(package.changes)
    edit(package)
    assert len(package.changes) == count + 1, "one ChangeReport a write"
    report = package.changes[-1]
    assert report.operation == operation
    assert report.addresses, "what was touched"
    assert "/word/document.xml" in report.parts_touched
    assert ends_with_paragraph(first_cell(package).element)
    if write == "Paragraph.delete" and not tracked:
        # the paragraph stays where it was, emptied: its text gone and no block shifted
        assert (report.text_before, report.text_after, report.moved) == ("CELL", "", ())
    if write == "insert_table End":
        # the cell's mark after the table is a new paragraph, its w14:paraId among the report's
        mark = first_cell(package).body.content[-1]
        assert mark.para_id and mark.para_id in report.created_para_ids
    load(package.save())


def test_the_two_functions_are_exported_as_core_ts_exports_them():
    from docx4j_py.model import content

    assert {"ends_with_paragraph", "end_with_paragraph"} <= set(content.__all__)
    assert re.search(r"TableRules", content.table_rules.__doc__ or "")
