"""Tracked changes inside a block-, paragraph-, row- and cell-level ``mc:AlternateContent``. CR-012 Phases A and A2.

The port of docx4j's ``TrackedChangesAlternateContentTest`` (``2853b7280``, its
CR-051 section 10.3 item 3 and 10.5 item 3; and ``9db1a3cbd``, its phase 2,
section 4.4 item 3 here): the changes of the branch Word reads are listed once,
each with the other branch's change of its kind at its position as a twin,
which resolves with it; ``accept_all``'s listing walks every branch and counts
the read branch's. A table left with no rows goes only when no branch has a row
left. The trees are built outside a part (``set_contents``), so the elements
stay whole whatever the load mode.

Before Phase A the engine read a block- or paragraph-level element as holding
nothing, and before Phase A2 a row- or cell-level one: no change in it was
listed, accepted or rejected.
"""

from __future__ import annotations

from docx4j_py import create_package
from docx4j_py.fragments import wml
from docx4j_py.model import revisions
from docx4j_py.model.revisions.lister import Lister

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
)


def ins(change_id: int, text: str) -> str:
    return f'<w:ins w:id="{change_id}" w:author="A"><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:ins>'


def deleted(change_id: int, text: str) -> str:
    return f'<w:del w:id="{change_id}" w:author="A"><w:r><w:delText xml:space="preserve">{text}</w:delText></w:r></w:del>'


#: A block-level element, an insertion and a deletion in each branch's paragraph.
BLOCK = (
    "<w:p><w:r><w:t>before</w:t></w:r></w:p>"
    f'<mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>C </w:t></w:r>{ins(1, "INSERTED")}{deleted(2, "DELETED")}</w:p></mc:Choice>'
    f"<mc:Fallback><w:p><w:r><w:t>F </w:t></w:r>{ins(3, 'INSERTED')}{deleted(4, 'DELETED')}</w:p></mc:Fallback></mc:AlternateContent>"
    "<w:p><w:r><w:t>after</w:t></w:r></w:p>"
)

#: A paragraph-level element, the same changes in each branch.
PARA = (
    "<w:p><w:r><w:t>x </w:t></w:r>"
    f'<mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>C </w:t></w:r>{ins(1, "INSERTED")}{deleted(2, "DELETED")}</mc:Choice>'
    f"<mc:Fallback><w:r><w:t>F </w:t></w:r>{ins(3, 'INSERTED')}{deleted(4, 'DELETED')}</mc:Fallback></mc:AlternateContent></w:p>"
)


TBL_PROPS = '<w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>'

#: A row-level element among a table's rows (docx4j CR-051 phase 2), the same changes in each branch's row.
ROW = (
    f"<w:tbl>{TBL_PROPS}"
    f'<mc:AlternateContent><mc:Choice Requires="w14"><w:tr><w:tc><w:p><w:r><w:t>C </w:t></w:r>{ins(1, "INSERTED")}{deleted(2, "DELETED")}</w:p></w:tc></w:tr></mc:Choice>'
    f"<mc:Fallback><w:tr><w:tc><w:p><w:r><w:t>F </w:t></w:r>{ins(3, 'INSERTED')}{deleted(4, 'DELETED')}</w:p></w:tc></w:tr></mc:Fallback></mc:AlternateContent>"
    "</w:tbl>"
)

#: A cell-level element among a row's cells, the same changes in each branch's cell.
CELL = (
    f"<w:tbl>{TBL_PROPS}<w:tr>"
    f'<mc:AlternateContent><mc:Choice Requires="w14"><w:tc><w:p><w:r><w:t>C </w:t></w:r>{ins(1, "INSERTED")}{deleted(2, "DELETED")}</w:p></w:tc></mc:Choice>'
    f"<mc:Fallback><w:tc><w:p><w:r><w:t>F </w:t></w:r>{ins(3, 'INSERTED')}{deleted(4, 'DELETED')}</w:p></w:tc></mc:Fallback></mc:AlternateContent>"
    "</w:tr></w:tbl>"
)


def deleted_row(change_id: int, text: str) -> str:
    return (
        f'<w:tr><w:trPr><w:del w:id="{change_id}" w:author="A"/></w:trPr>'
        f"<w:tc><w:p><w:r><w:delText>{text}</w:delText></w:r></w:p></w:tc></w:tr>"
    )


def plain_row(text: str) -> str:
    return f"<w:tr><w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:tr>"


def with_body(body: str):
    pkg = create_package()
    pkg.main_document_part.set_contents(
        wml(f"<w:document {NS}><w:body>{body}</w:body></w:document>")
    )
    return pkg


def body_xml(pkg) -> str:
    return pkg.main_document_part.get_xml()


def listed_once_with_twins(body: str) -> None:
    pkg = with_body(body)
    changes = revisions.tracked_changes(pkg)
    assert len(changes) == 2, changes
    for change in changes:
        assert len(change.twins) == 1, change
        assert change.twins[0].kind == change.kind
    assert "INSERTED" in changes[0].text
    assert "DELETED" in changes[1].text
    # every branch's changes are listed for accept_all
    assert len(Lister(pkg, False).list()) == 4


def test_block_level_changes_are_listed_once_with_twins() -> None:
    listed_once_with_twins(BLOCK)


def test_paragraph_level_changes_are_listed_once_with_twins() -> None:
    listed_once_with_twins(PARA)


def accepting_resolves_the_twin(body: str) -> None:
    pkg = with_body(body)
    revisions.tracked_changes(pkg)[0].accept()
    xml = body_xml(pkg)
    assert "w:ins " not in xml
    assert xml.count("INSERTED") == 2, "the insertion stays in both branches"
    assert len(revisions.tracked_changes(pkg)) == 1, (
        "the deletions are still listed, once"
    )
    revisions.tracked_changes(pkg)[0].reject()
    xml = body_xml(pkg)
    assert "w:del " not in xml
    assert xml.count("DELETED") == 2, "the deleted text is back in both branches"
    assert revisions.tracked_changes(pkg) == []
    assert Lister(pkg, False).list() == []


def test_accepting_a_block_level_change_resolves_its_twin() -> None:
    accepting_resolves_the_twin(BLOCK)


def test_accepting_a_paragraph_level_change_resolves_its_twin() -> None:
    accepting_resolves_the_twin(PARA)


def accept_all_resolves_every_branch(body: str) -> None:
    pkg = with_body(body)
    assert revisions.accept_all(pkg) == 2
    xml = body_xml(pkg)
    assert "w:ins " not in xml and "w:del " not in xml
    assert xml.count("INSERTED") == 2
    assert xml.count("DELETED") == 0
    assert "<mc:Choice" in xml and "<mc:Fallback" in xml, "the element is still whole"


def test_accept_all_resolves_every_branch_at_block_level() -> None:
    accept_all_resolves_every_branch(BLOCK)


def test_accept_all_resolves_every_branch_at_paragraph_level() -> None:
    accept_all_resolves_every_branch(PARA)


def test_the_accepted_view_reads_the_branch_word_reads() -> None:
    """The text model's view of a cell reads the branch Word reads (docx4j: Texts.cellText through Trees.blockItems)."""
    pkg = with_body(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>'
        + BLOCK
        + "</w:tc></w:tr></w:tbl>"
    )
    change = revisions.tracked_changes(pkg)[0]
    assert "INSERTED" in change.text
    assert len(change.twins) == 1


def test_reject_all_resolves_every_branch() -> None:
    """The mirror of accept_all, beyond docx4j's test: both branches back as they were before the changes."""
    for body in (BLOCK, PARA):
        pkg = with_body(body)
        assert revisions.reject_all(pkg) == 2
        xml = body_xml(pkg)
        assert "w:ins " not in xml and "w:del " not in xml
        assert xml.count("INSERTED") == 0
        assert xml.count("DELETED") == 2


def test_a_text_box_in_the_other_branch_is_the_twin_of_the_read_branchs() -> None:
    """docx4j CR-051 section 11.2: text boxes in another branch become twins of the listed branch's by position."""
    box = (
        '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        '<wp:extent cx="1" cy="1"/><wp:docPr id="1" name="b"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr/>'
        "<wps:txbx><w:txbxContent><w:p>{}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>"
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )
    body = (
        f'<mc:AlternateContent><mc:Choice Requires="w14">{box.format(ins(1, "BOXED"))}</mc:Choice>'
        f"<mc:Fallback>{box.format(ins(2, 'BOXED'))}</mc:Fallback></mc:AlternateContent>"
    )
    pkg = with_body(body)
    (change,) = revisions.tracked_changes(pkg)
    assert "BOXED" in change.text
    assert len(change.twins) == 1
    change.accept()
    xml = body_xml(pkg)
    assert "w:ins " not in xml and xml.count("BOXED") == 2


# ---------------------------------------------------------------------------
# row and cell level (CR-012 Phase A2, docx4j CR-051 phase 2 at 9db1a3cbd)
# ---------------------------------------------------------------------------


def test_row_level_changes_are_listed_once_with_twins() -> None:
    listed_once_with_twins(ROW)


def test_cell_level_changes_are_listed_once_with_twins() -> None:
    listed_once_with_twins(CELL)


def test_accepting_a_row_level_change_resolves_its_twin() -> None:
    accepting_resolves_the_twin(ROW)


def test_accepting_a_cell_level_change_resolves_its_twin() -> None:
    accepting_resolves_the_twin(CELL)


def test_accept_all_resolves_every_branch_at_row_level() -> None:
    accept_all_resolves_every_branch(ROW)


def test_accept_all_resolves_every_branch_at_cell_level() -> None:
    accept_all_resolves_every_branch(CELL)


def test_accepting_a_deleted_row_in_every_branch_removes_the_table() -> None:
    """Listed once with its twin; accepted, both rows go, and the table with them, no branch having a row left."""
    pkg = with_body(
        f"<w:p><w:r><w:t>before</w:t></w:r></w:p><w:tbl>{TBL_PROPS}"
        f'<mc:AlternateContent><mc:Choice Requires="w14">{deleted_row(1, "C")}</mc:Choice>'
        f"<mc:Fallback>{deleted_row(2, 'F')}</mc:Fallback></mc:AlternateContent></w:tbl>"
    )
    (change,) = revisions.tracked_changes(pkg)
    assert len(change.twins) == 1
    change.accept()
    xml = body_xml(pkg)
    assert "<w:tbl" not in xml and "<w:tr" not in xml
    assert "before" in xml


def test_accepting_a_deleted_row_keeps_a_table_whose_other_branch_has_a_row() -> None:
    """A deleted row in the branch Word reads, a plain row in the Fallback: accepted, the table stays."""
    pkg = with_body(
        f"<w:tbl>{TBL_PROPS}"
        f'<mc:AlternateContent><mc:Choice Requires="w14">{deleted_row(1, "C")}</mc:Choice>'
        f"<mc:Fallback>{plain_row('F')}</mc:Fallback></mc:AlternateContent></w:tbl>"
    )
    (change,) = revisions.tracked_changes(pkg)
    assert change.twins == []
    change.accept()
    xml = body_xml(pkg)
    assert "<w:tbl" in xml and "<mc:Fallback>" in xml and "<w:tr" in xml
    assert '<mc:Choice Requires="w14"><w:tr' not in xml
    assert ">F<" in xml


def test_a_rows_text_reads_through_a_cell_level_element() -> None:
    """docx4j Texts.rowText over Lister.cellsOf (its CR-051 phase 2): a deleted row's text has Word's branch's cell."""
    pkg = with_body(
        f'<w:tbl>{TBL_PROPS}<w:tr><w:trPr><w:del w:id="1" w:author="A"/></w:trPr>'
        "<w:tc><w:p><w:r><w:delText>A</w:delText></w:r></w:p></w:tc>"
        '<mc:AlternateContent><mc:Choice Requires="w14"><w:tc><w:p><w:r><w:delText>C</w:delText></w:r></w:p></w:tc></mc:Choice>'
        "<mc:Fallback><w:tc><w:p><w:r><w:delText>F</w:delText></w:r></w:p></w:tc></mc:Fallback></mc:AlternateContent>"
        "</w:tr></w:tbl>"
    )
    (change,) = revisions.tracked_changes(pkg)
    assert change.text == "A\tC\r\n"


def test_a_nested_tables_text_reads_through_a_row_level_element() -> None:
    """Texts.cellText's nested table through Lister.rowsOf: Word's branch's row, not passed by."""
    nested = (
        f"<w:tbl>{TBL_PROPS}"
        f'<mc:AlternateContent><mc:Choice Requires="w14">{plain_row("NC")}</mc:Choice>'
        f"<mc:Fallback>{plain_row('NF')}</mc:Fallback></mc:AlternateContent></w:tbl>"
    )
    pkg = with_body(
        f'<w:tbl>{TBL_PROPS}<w:tr><w:trPr><w:ins w:id="1" w:author="A"/></w:trPr>'
        f"<w:tc>{nested}<w:p/></w:tc></w:tr></w:tbl>"
    )
    (change,) = revisions.tracked_changes(pkg)
    assert change.text == "NC\r\n\r\n"
