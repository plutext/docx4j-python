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

A cell's twin, a ``w:cellDel`` accepted or a ``w:cellIns`` rejected in a branch
Word does not read, is removed from its own branch, and that branch's neighbour
widened (docx4j ``60bd4b9ad``, its CR-051 section 12.5; CR-012 section 10.10).

A table's property record, a ``w:tcPrChange`` or ``w:trPrChange`` in a branch
Word does not read, is part of the table's one change: listed, and resolved in
every branch, and a merged grid collapsed along each branch's reading of a row
(docx4j ``901a922a3``, its CR-051 section 12.5; CR-012 section 10.11). Every
layout's spans are counted from the old ones before any is written, and the
change's text is the rows of the branch Word reads (docx4j ``84b202c5a``;
CR-012 section 10.12).
"""

from __future__ import annotations

from collections.abc import Callable

from docx4j_py import create_package, load
from docx4j_py.child import is_alternate_content
from docx4j_py.fragments import wml
from docx4j_py.model import revisions
from docx4j_py.model.revisions import tables
from docx4j_py.model.revisions.lister import Lister
from docx4j_py.openpackaging import LoadOptions

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


# ---------------------------------------------------------------------------
# a cell's twin in its own branch (docx4j 60bd4b9ad, its CR-051 section 12.5)
# ---------------------------------------------------------------------------


def cell(text: str, tc_pr: str) -> str:
    return f"<w:tc><w:tcPr>{tc_pr}</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"


def deleted_cell(change_id: int, text: str) -> str:
    return cell(text, f'<w:tcW w:w="1000" w:type="dxa"/><w:cellDel w:id="{change_id}" w:author="A"/>')


def inserted_cell(change_id: int, text: str) -> str:
    """As :func:`deleted_cell`, a ``w:cellIns``: beyond docx4j's test, whose commit names the case."""
    return cell(text, f'<w:tcW w:w="1000" w:type="dxa"/><w:cellIns w:id="{change_id}" w:author="A"/>')


def plain_cell(text: str) -> str:
    return cell(text, '<w:tcW w:w="1000" w:type="dxa"/>')


Marked = Callable[[int, str], str]


def twin_beside(marked: Marked) -> str:
    """docx4j's first shape: a marked cell in each branch, each branch with a plain cell beside it."""
    return (
        f"<w:tbl>{TBL_PROPS}<w:tr>"
        f'<mc:AlternateContent><mc:Choice Requires="w14">{marked(1, "C")}{plain_cell("Ckeep")}</mc:Choice>'
        f"<mc:Fallback>{marked(2, 'F')}{plain_cell('Fkeep')}</mc:Fallback></mc:AlternateContent></w:tr></w:tbl>"
    )


def twin_alone(marked: Marked) -> str:
    """docx4j's second shape: the marked cell alone in each branch, a plain cell beside the element in the row."""
    return (
        f"<w:tbl>{TBL_PROPS}<w:tr>{plain_cell('keep')}"
        f'<mc:AlternateContent><mc:Choice Requires="w14">{marked(1, "C")}</mc:Choice>'
        f"<mc:Fallback>{marked(2, 'F')}</mc:Fallback></mc:AlternateContent></w:tr></w:tbl>"
    )


def first_row(pkg):
    return pkg.main_document_part.contents.body.content[0].content[0]


def branches_of_the_element(pkg) -> list:
    """The branches of the first row's cell-level element, its Choice first."""
    ac = next(item for item in first_row(pkg).content if is_alternate_content(item))
    return [*ac.choice, ac.fallback]


def gone_from_both_branches(xml: str) -> None:
    assert "cellDel" not in xml and "cellIns" not in xml, xml
    assert ">C<" not in xml and ">F<" not in xml, xml


def each_branchs_neighbour_widened_once(pkg) -> None:
    """docx4j's first test's assertions after the resolution."""
    xml = body_xml(pkg)
    gone_from_both_branches(xml)
    assert "Ckeep" in xml and "Fkeep" in xml
    assert xml.count('w:gridSpan w:val="2"') == 2, "each branch's neighbour widened once"
    assert xml.count('w:tcW w:w="2000"') == 2
    assert revisions.tracked_changes(pkg) == []
    assert Lister(pkg, False).list() == []


def the_rows_neighbour_widened_once(pkg) -> None:
    """docx4j's second test's assertions after the resolution."""
    xml = body_xml(pkg)
    gone_from_both_branches(xml)
    assert xml.count('w:gridSpan w:val="2"') == 1, "the row's neighbour widened once"
    assert xml.count('w:tcW w:w="2000"') == 1
    assert "<mc:Choice" in xml and "<mc:Fallback" in xml, "the element stays"
    assert all(tables.cells_in(branch) == [] for branch in branches_of_the_element(pkg)), (
        "its branches now empty of cells"
    )
    assert revisions.tracked_changes(pkg) == []


def listed_once_with_its_twin(pkg):
    (change,) = revisions.tracked_changes(pkg)
    assert len(change.twins) == 1, change
    assert change.twins[0].kind == change.kind
    return change


def test_accepting_a_deleted_cell_resolves_its_twin_in_its_own_branch() -> None:
    """docx4j ``acceptingADeletedCellResolvesItsTwinInItsOwnBranch`` (this package's finding, CR-012 section 10.8).

    A deleted cell in each branch of a cell-level element, each branch with a
    plain cell beside it: accepted, both branches lose their cell and each
    branch's own neighbour is widened, once. Before ``60bd4b9ad`` the twin kept
    its cell and its marker.
    """
    pkg = with_body(twin_beside(deleted_cell))
    listed_once_with_its_twin(pkg).accept()
    each_branchs_neighbour_widened_once(pkg)


def test_accepting_a_deleted_cell_alone_in_each_branch_widens_the_rows_neighbour_once() -> None:
    """docx4j ``acceptingADeletedCellAloneInEachBranchWidensTheRowsNeighbourOnce``.

    Each branch holding the deleted cell alone, a plain cell beside the element
    in the row: the row's neighbour is widened once, not once per branch.
    """
    pkg = with_body(twin_alone(deleted_cell))
    listed_once_with_its_twin(pkg).accept()
    the_rows_neighbour_widened_once(pkg)


def test_rejecting_an_inserted_cell_resolves_its_twin_in_its_own_branch() -> None:
    """A ``w:cellIns`` rejected, the case ``60bd4b9ad``'s message names and docx4j does not test: as the deleted cell accepted."""
    pkg = with_body(twin_beside(inserted_cell))
    listed_once_with_its_twin(pkg).reject()
    each_branchs_neighbour_widened_once(pkg)


def test_rejecting_an_inserted_cell_alone_in_each_branch_widens_the_rows_neighbour_once() -> None:
    pkg = with_body(twin_alone(inserted_cell))
    listed_once_with_its_twin(pkg).reject()
    the_rows_neighbour_widened_once(pkg)


def test_accept_all_resolves_a_cells_twin_in_its_own_branch() -> None:
    """``accept_all`` resolves every branch's, last first, so the twin goes before the read branch's cell.

    The counts are the single change's either way; an inserted cell accepted
    keeps its cell in both branches, its marker gone, nothing widened.
    """
    pkg = with_body(twin_beside(deleted_cell))
    assert revisions.accept_all(pkg) == 1
    each_branchs_neighbour_widened_once(pkg)

    pkg = with_body(twin_alone(deleted_cell))
    assert revisions.accept_all(pkg) == 1
    the_rows_neighbour_widened_once(pkg)

    pkg = with_body(twin_beside(inserted_cell))
    assert revisions.accept_all(pkg) == 1
    xml = body_xml(pkg)
    assert "cellIns" not in xml
    assert all(text in xml for text in (">C<", "Ckeep", ">F<", "Fkeep"))
    assert "w:gridSpan" not in xml and 'w:tcW w:w="2000"' not in xml
    assert Lister(pkg, False).list() == []


def test_reject_all_resolves_a_cells_twin_in_its_own_branch() -> None:
    """The mirror of :func:`test_accept_all_resolves_a_cells_twin_in_its_own_branch`."""
    pkg = with_body(twin_beside(inserted_cell))
    assert revisions.reject_all(pkg) == 1
    each_branchs_neighbour_widened_once(pkg)

    pkg = with_body(twin_alone(inserted_cell))
    assert revisions.reject_all(pkg) == 1
    the_rows_neighbour_widened_once(pkg)

    pkg = with_body(twin_beside(deleted_cell))
    assert revisions.reject_all(pkg) == 1
    xml = body_xml(pkg)
    assert "cellDel" not in xml
    assert all(text in xml for text in (">C<", "Ckeep", ">F<", "Fkeep"))
    assert "w:gridSpan" not in xml and 'w:tcW w:w="2000"' not in xml
    assert Lister(pkg, False).list() == []


def _span(tc) -> int:
    grid_span = tc.tc_pr.grid_span if tc.tc_pr is not None else None
    return int(grid_span.val) if grid_span is not None and grid_span.val is not None else 1


def columns_through_each_branch(pkg) -> list[int]:
    """The grid columns the first row's cells span, read through each branch of its element in turn, Choice first."""
    tr = first_row(pkg)
    out = []
    for branch in branches_of_the_element(pkg):
        total = 0
        for item in tr.content:
            if is_alternate_content(item):
                total += sum(_span(tc) for tc, _owner in tables.cells_in(branch))
            else:
                total += _span(item)
        out.append(total)
    return out


def test_known_gap_pinned_a_lone_twin_beside_a_read_branch_of_two_cells_leaves_its_branch_short() -> None:
    """PINNED KNOWN GAP, not wanted behaviour: docx4j CR-051 section 12.5 leaves it, and so does this port.

    A read branch of two cells beside a twin branch of one, in a shape no
    producer writes (Word writes no cell-level element at all, docx4j CR-051
    phase 0). The Choice holds a deleted cell and a plain one; the Fallback
    holds only the deleted cell's twin, spanning the same two grid columns; a
    plain cell follows the element. Accepted, the read branch's cell goes and
    its neighbour in that branch (``Ckeep``) is widened. The twin, alone in its
    branch, goes with nothing widened, on the rule that the row's neighbour was
    widened already. Here it was not. So the Fallback's reading of the row is
    short by the twin's span, one ``w:gridSpan`` of two columns: 3 and 3
    before, 3 and 1 after.

    A fix would widen the row's neighbour for the twin only when the primary's
    neighbour was in its own branch. If this test fails, the gap has moved:
    update CR-012 section 10.10 with docx4j.
    """
    twin = cell("F", '<w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="2"/><w:cellDel w:id="2" w:author="A"/>')
    pkg = with_body(
        f"<w:tbl>{TBL_PROPS}<w:tr>"
        f'<mc:AlternateContent><mc:Choice Requires="w14">{deleted_cell(1, "C")}{plain_cell("Ckeep")}</mc:Choice>'
        f"<mc:Fallback>{twin}</mc:Fallback></mc:AlternateContent>{plain_cell('keep')}</w:tr></w:tbl>"
    )
    assert columns_through_each_branch(pkg) == [3, 3], "every branch spans the same grid columns"
    listed_once_with_its_twin(pkg).accept()
    xml = body_xml(pkg)
    gone_from_both_branches(xml)
    assert xml.count('w:gridSpan w:val="2"') == 1 and xml.count('w:tcW w:w="2000"') == 1, "Ckeep's alone"
    choice, fallback = branches_of_the_element(pkg)
    ((ckeep, _owner),) = tables.cells_in(choice)
    assert _span(ckeep) == 2
    assert tables.cells_in(fallback) == []
    keep = first_row(pkg).content[-1]
    assert _span(keep) == 1 and keep.tc_pr.tc_w.w == 1000, "the row's neighbour is not widened"
    assert columns_through_each_branch(pkg) == [3, 1], "the known gap: the Fallback's row is the twin's span short"


def test_word_mode_has_no_twin_to_resolve() -> None:
    """In ``word`` mode the element is replaced on load by Word's branch, so there is no twin: one change, one widening."""
    pkg = load(with_body(twin_beside(deleted_cell)).save(), options=LoadOptions(mc_block_level="word"))
    (change,) = revisions.tracked_changes(pkg)
    assert change.twins == []
    assert len(pkg.main_document_part.resolved_alternate_content) == 1, "recorded when the part was read"
    change.accept()
    xml = body_xml(pkg)
    assert "AlternateContent" not in xml and "cellDel" not in xml
    assert ">C<" not in xml and "Ckeep" in xml
    assert "F<" not in xml and "Fkeep" not in xml, "the Fallback went on load"
    assert xml.count('w:gridSpan w:val="2"') == 1
    assert xml.count('w:tcW w:w="2000"') == 1
    assert revisions.tracked_changes(pkg) == []


# ---------------------------------------------------------------------------
# a table's property records in every branch (docx4j 901a922a3, its CR-051 section 12.5)
# ---------------------------------------------------------------------------


def changed_cell(change_id: int, text: str) -> str:
    """docx4j's ``changedCell``: a red-shaded cell whose ``w:tcPrChange`` records it unshaded."""
    return cell(
        text,
        '<w:tcW w:w="1000" w:type="dxa"/><w:shd w:val="clear" w:fill="FF0000"/>'
        f'<w:tcPrChange w:id="{change_id}" w:author="A"><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr></w:tcPrChange>',
    )


def cell_record_in_each_branch(choice: str = "C", fallback: str = "F") -> str:
    """docx4j's first shape: a ``w:tcPrChange`` in each branch's cell of a cell-level element."""
    return (
        f"<w:tbl>{TBL_PROPS}<w:tr>"
        f'<mc:AlternateContent><mc:Choice Requires="w14">{choice and changed_cell(1, choice)}</mc:Choice>'
        f"<mc:Fallback>{fallback and changed_cell(2, fallback)}</mc:Fallback></mc:AlternateContent></w:tr></w:tbl>"
    )


#: docx4j's second shape's row: a ``w:trPrChange`` recording no properties, the row now ``w:cantSplit``.
CHANGED_ROW = (
    '<w:tr><w:trPr><w:cantSplit/><w:trPrChange w:id="1" w:author="A"><w:trPr/></w:trPrChange></w:trPr>'
    f"{plain_cell('F')}</w:tr>"
)

#: docx4j's second shape: the record only on the Fallback's row of a row-level element.
ROW_RECORD_IN_THE_FALLBACK = (
    f"<w:tbl>{TBL_PROPS}"
    f'<mc:AlternateContent><mc:Choice Requires="w14">{plain_row("C")}</mc:Choice>'
    f"<mc:Fallback>{CHANGED_ROW}</mc:Fallback></mc:AlternateContent></w:tbl>"
)


def first_table(pkg):
    return pkg.main_document_part.contents.body.content[0]


def both_cells_accepted(pkg) -> None:
    """docx4j's first test's assertions after accepting: every branch's record gone, its shading kept."""
    xml = body_xml(pkg)
    assert xml.count("tcPrChange") == 0, xml
    assert xml.count('w:fill="FF0000"') == 2, "accepted: both branches keep the shading"
    assert revisions.tracked_changes(pkg) == []
    assert Lister(pkg, False).list() == []


def both_cells_rejected(pkg) -> None:
    """docx4j's first test's assertions after rejecting: every branch's cell gets its recorded properties back."""
    xml = body_xml(pkg)
    assert xml.count("tcPrChange") == 0, xml
    assert xml.count('w:fill="FF0000"') == 0, "rejected: both branches lose the shading"
    assert xml.count('<w:tcW w:w="1000"') == 2
    assert revisions.tracked_changes(pkg) == []
    assert Lister(pkg, False).list() == []


def test_resolving_a_table_property_record_reaches_every_branchs_cell() -> None:
    """docx4j ``resolvingATablePropertyRecordReachesEveryBranchsCell`` (this package's finding, CR-012 section 10.10).

    A ``w:tcPrChange`` in each branch's cell of a cell-level element: the
    table's one property record is listed once; accepted, every branch's record
    goes; rejected, every branch's cell gets its recorded properties back.
    Before ``901a922a3`` the Fallback's record was kept either way.
    """
    pkg = with_body(cell_record_in_each_branch())
    assert len(revisions.tracked_changes(pkg)) == 1, revisions.tracked_changes(pkg)
    assert revisions.accept_all(pkg) == 1
    both_cells_accepted(pkg)

    pkg = with_body(cell_record_in_each_branch())
    assert revisions.reject_all(pkg) == 1
    both_cells_rejected(pkg)


def test_accepting_a_table_property_record_reaches_a_row_in_another_branch() -> None:
    """docx4j ``acceptingATablePropertyRecordReachesARowInAnotherBranch``.

    A ``w:trPrChange`` only the Fallback's row of a row-level element holds:
    listed, and accepted, the record goes and the row keeps ``w:cantSplit``.
    Before ``901a922a3`` it was not listed, and ``accept_all`` kept it.
    """
    pkg = with_body(ROW_RECORD_IN_THE_FALLBACK)
    assert len(revisions.tracked_changes(pkg)) == 1, revisions.tracked_changes(pkg)
    assert revisions.accept_all(pkg) == 1
    xml = body_xml(pkg)
    assert xml.count("trPrChange") == 0, xml
    assert "<w:cantSplit/>" in xml
    assert revisions.tracked_changes(pkg) == []


def test_the_listed_table_change_accepted_or_rejected_reaches_every_branchs_cell() -> None:
    """Beyond docx4j's test: the one listed change's own ``accept()`` and ``reject()``, not only ``*_all``.

    The table's change has no twin (its records are one change, read through
    every branch), and its text is the row's as Word reads it, the Choice's cell.
    """
    pkg = with_body(cell_record_in_each_branch())
    (change,) = revisions.tracked_changes(pkg)
    assert change.twins == []
    assert change.text == "C\r\n"
    change.accept()
    both_cells_accepted(pkg)

    pkg = with_body(cell_record_in_each_branch())
    (change,) = revisions.tracked_changes(pkg)
    change.reject()
    both_cells_rejected(pkg)


def test_rejecting_a_table_property_record_reaches_a_row_in_another_branch() -> None:
    """Beyond docx4j's test: the Fallback's ``w:trPrChange`` rejected, by its change and by ``reject_all``.

    The record held no properties, so the row's ``w:trPr`` goes whole, its
    ``w:cantSplit`` with it; the Choice's row is left as it was. The change has
    no text: its text is the rows of the branch Word reads, and the Choice's
    row is not in the change (docx4j ``changeText`` on ``Lister.rowsOf`` since
    ``84b202c5a``; ``901a922a3`` gave the Fallback's row, ``F\\r\\n``).
    """
    for reject in (lambda pkg: revisions.tracked_changes(pkg)[0].reject(), revisions.reject_all):
        pkg = with_body(ROW_RECORD_IN_THE_FALLBACK)
        (change,) = revisions.tracked_changes(pkg)
        assert change.text == "", "the rows Word reads: none is in the change"
        reject(pkg)
        xml = body_xml(pkg)
        assert "trPrChange" not in xml and "cantSplit" not in xml and "<w:trPr" not in xml, xml
        assert '<mc:Choice Requires="w14"><w:tr><w:tc><w:p><w:r><w:t>C</w:t>' in xml
        assert ">F<" in xml
        assert revisions.tracked_changes(pkg) == []
        assert Lister(pkg, False).list() == []


def test_a_record_only_a_branch_word_does_not_read_holds_is_one_of_the_tables() -> None:
    """Beyond docx4j's test: ``has_records``, ``in_change`` and ``record_of`` answer for it.

    A ``w:tcPrChange`` only in the Fallback's cell of a cell-level element, and
    a ``w:trPrChange`` only on the Fallback's row of a row-level element. Before
    ``901a922a3`` the table had no records, the row was not in the change, and
    nothing was listed.
    """
    pkg = with_body(cell_record_in_each_branch(choice=""))
    tbl, tr = first_table(pkg), first_row(pkg)
    assert tables.in_change(tr)
    assert tables.has_records(tbl)
    fallback = branches_of_the_element(pkg)[1]
    ((tc, _owner),) = tables.cells_in(fallback)
    assert tables.record_of(tbl) is tc.tc_pr.tc_pr_change
    (change,) = revisions.tracked_changes(pkg)
    change.accept()
    xml = body_xml(pkg)
    assert "tcPrChange" not in xml and xml.count('w:fill="FF0000"') == 1
    assert not tables.has_records(tbl)

    pkg = with_body(ROW_RECORD_IN_THE_FALLBACK)
    tbl = first_table(pkg)
    assert tables.has_records(tbl)
    assert [tables.in_change(tr) for tr, _owner in tables.all_rows_of(tbl)] == [False, True]
    assert [tables.in_change(tr) for tr, _owner in tables.rows_of(tbl)] == [False], "Word's branch: the Choice's row"
    fallback_row = tbl.content[0].fallback.any_element[0]
    assert tables.record_of(tbl) is fallback_row.tr_pr.tr_pr_change


# -- a merged grid collapsed along every branch reading of a row ------------------


def spanned(text: str, span: int) -> str:
    """A cell over `span` grid columns, with a ``w:tcPr`` either way, so that its span is counted again."""
    return cell(text, f'<w:gridSpan w:val="{span}"/>' if span > 1 else '<w:tcW w:w="0" w:type="auto"/>')


def regridded(columns: int, rows: str) -> str:
    """A table whose recorded grid had `columns` columns of 1000, its ``w:tblPr`` changed too, so that it is listed.

    Rejected, the recorded grid comes back and is collapsed (check 27, C23):
    the columns no row keeps apart are merged, those past every row's last cell
    dropped.
    """
    grid = '<w:gridCol w:w="1000"/>' * columns
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblPrChange w:id="9" w:author="A"><w:tblPr/></w:tblPrChange></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="1000"/>'
        f'<w:tblGridChange w:id="10"><w:tblGrid>{grid}</w:tblGrid></w:tblGridChange></w:tblGrid>'
        f"{rows}</w:tbl>"
    )


def grid_of(pkg) -> list[int]:
    return [int(col.w) for col in first_table(pkg).tbl_grid.grid_col]


def spans_in(branch) -> list[int]:
    return [_span(tc) for tc, _owner in tables.cells_in(branch)]


def test_rejecting_a_merged_grid_counts_it_along_every_branch_reading_of_a_row() -> None:
    """``collapse_grid`` along ``cell_layouts_of`` (docx4j ``Tables.collapseGrid`` since ``901a922a3``).

    The recorded grid is six columns. The row skips two (``w:gridBefore``),
    then a cell-level element: the Choice's cells span 1 and 3 columns, the
    Fallback's 2 and 2, so the Choice keeps edge 3 apart and the Fallback edge
    4. Rejected, the edges of both readings are kept: four columns, 2000, 1000,
    1000 and 2000. ``w:gridBefore`` is read once and counted again (2 to 1),
    and each branch's spans are counted again along its own reading: the
    Choice's 1 and 2, the Fallback's 2 and 1. With Word's branch alone the
    grid was three columns and the Fallback kept spans of 2 and 2, five
    columns in a grid of three.
    """
    pkg = with_body(
        regridded(
            6,
            '<w:tr><w:trPr><w:gridBefore w:val="2"/></w:trPr>'
            f'<mc:AlternateContent><mc:Choice Requires="w14">{spanned("C1", 1)}{spanned("C2", 3)}</mc:Choice>'
            f'<mc:Fallback>{spanned("F1", 2)}{spanned("F2", 2)}</mc:Fallback></mc:AlternateContent></w:tr>',
        )
    )
    (change,) = revisions.tracked_changes(pkg)
    change.reject()
    assert grid_of(pkg) == [2000, 1000, 1000, 2000]
    tr = first_row(pkg)
    assert tr.tr_pr.content[0].val == 1, "w:gridBefore counted again, once"
    choice, fallback = branches_of_the_element(pkg)
    assert spans_in(choice) == [1, 2]
    assert spans_in(fallback) == [2, 1]
    assert columns_through_each_branch(pkg) == [3, 3], "each reading of the row spans the grid, less w:gridBefore"
    assert revisions.tracked_changes(pkg) == []


def test_a_collapsed_rows_grid_after_is_counted_from_the_branch_word_reads() -> None:
    """docx4j's rule: the row's ``w:gridAfter`` from the branch Word reads, the first layout.

    A shape no producer writes, branches ending at different columns, so that
    the rule decides: the recorded grid is four columns; the row's Choice cell
    spans 3, its Fallback cell 2, and the row skips one column after its cells.
    Rejected, the grid is two columns (2000 and 1000), the fourth past every
    row's last cell dropped. Word's branch ends on the last edge kept, so the
    row's ``w:gridAfter`` and ``w:wAfter`` go; counted from the Fallback's end
    it would stay 1. The Choice's cell spans 2, the Fallback's 1.
    """
    pkg = with_body(
        regridded(
            4,
            '<w:tr><w:trPr><w:gridAfter w:val="1"/><w:wAfter w:w="1000" w:type="dxa"/></w:trPr>'
            f'<mc:AlternateContent><mc:Choice Requires="w14">{spanned("C", 3)}</mc:Choice>'
            f'<mc:Fallback>{spanned("F", 2)}</mc:Fallback></mc:AlternateContent></w:tr>',
        )
    )
    assert revisions.reject_all(pkg) == 1
    assert grid_of(pkg) == [2000, 1000]
    assert first_row(pkg).tr_pr is None, "w:gridAfter and w:wAfter gone, the w:trPr left empty pruned"
    choice, fallback = branches_of_the_element(pkg)
    assert spans_in(choice) == [2]
    assert spans_in(fallback) == [1]


def test_rejecting_a_merged_grid_counts_every_branchs_row() -> None:
    """``collapse_grid`` over ``all_rows_of``: a row-level element's Fallback row keeps its edges, and its spans are counted again.

    The recorded grid is four columns; the Choice's row spans 1 and 3, the
    Fallback's 2 and 2. Rejected, three columns (1000, 1000, 2000), the
    Choice's row 1 and 2, the Fallback's 2 and 1.
    """
    pkg = with_body(
        regridded(
            4,
            f'<mc:AlternateContent><mc:Choice Requires="w14"><w:tr>{spanned("A", 1)}{spanned("B", 3)}</w:tr></mc:Choice>'
            f'<mc:Fallback><w:tr>{spanned("X", 2)}{spanned("Y", 2)}</w:tr></mc:Fallback></mc:AlternateContent>',
        )
    )
    (change,) = revisions.tracked_changes(pkg)
    change.reject()
    assert grid_of(pkg) == [1000, 1000, 2000]
    ac = first_table(pkg).content[0]
    choice_row, fallback_row = ac.choice[0].any_element[0], ac.fallback.any_element[0]
    assert [_span(tc) for tc in choice_row.content] == [1, 2]
    assert [_span(tc) for tc in fallback_row.content] == [2, 1]


# -- every layout's spans counted from the old ones (docx4j 84b202c5a) -------------

#: docx4j's two tests' ``w:tblPr``: centred, recorded with no properties, so that the table's change is listed, and whole.
CHANGED_TBL_PR = (
    '<w:tblPr><w:jc w:val="center"/><w:tblPrChange w:id="1" w:author="A"><w:tblPr/></w:tblPrChange></w:tblPr>'
)


def spanned_in_500s(text: str, span: int) -> str:
    """docx4j's ``spanned`` (``84b202c5a``): a cell over `span` grid columns of 500 twips, its ``w:tcW`` theirs."""
    return cell(text, f'<w:tcW w:w="{500 * span}" w:type="dxa"/><w:gridSpan w:val="{span}"/>')


def regridded_in_500s(rows: str) -> str:
    """docx4j's grid change (``84b202c5a``): three columns of 1000 now, six of 500 recorded."""
    current = '<w:gridCol w:w="1000"/>' * 3
    recorded = '<w:gridCol w:w="500"/>' * 6
    return (
        f"<w:tbl>{CHANGED_TBL_PR}"
        f'<w:tblGrid>{current}<w:tblGridChange w:id="2"><w:tblGrid>{recorded}</w:tblGrid></w:tblGridChange></w:tblGrid>'
        f"{rows}</w:tbl>"
    )


def element(choice: str, fallback: str) -> str:
    return f'<mc:AlternateContent><mc:Choice Requires="w14">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>'


#: docx4j's shape (core-ts CR-013 phase 3's): row 1 X over two old columns, then
#: a cell-level element, C | F over four; row 2 a, b and c over two each.
GRID_CHANGE_X_BEFORE = regridded_in_500s(
    f'<w:tr>{spanned_in_500s("X", 2)}{element(spanned_in_500s("C", 4), spanned_in_500s("F", 4))}</w:tr>'
    f'<w:tr>{spanned_in_500s("a", 2)}{spanned_in_500s("b", 2)}{spanned_in_500s("c", 2)}</w:tr>'
)


def test_rejecting_a_grid_change_counts_every_branchs_spans_from_the_old_ones() -> None:
    """docx4j ``rejectingAGridChangeCountsEveryBranchsSpansFromTheOldOnes`` (core-ts's finding, its CR-013 phase 3).

    A cell outside the element is in every branch's layout, so the new spans
    are counted from the old ones along every layout before any is written.
    Restored grid 6 x 500, current 3 x 1000; rejected, the grid is 3 x 1000, X,
    a, b and c span 1 (no ``w:gridSpan``), C and F span 2. Before
    ``84b202c5a`` the Fallback's layout, counted first, wrote X's span 1, and
    Word's then placed C from column 1 where it starts at old column 2: C got
    span 1, the row two columns through the Choice in a grid of three.
    """
    pkg = with_body(GRID_CHANGE_X_BEFORE)
    assert revisions.reject_all(pkg) == 1
    xml = body_xml(pkg)
    assert xml.count("<w:gridCol ") == 3, xml
    assert xml.count('<w:gridCol w:w="1000"/>') == 3, xml
    assert xml.count("<w:gridSpan ") == 2, "C and F span 2, the others none"
    assert xml.count('<w:gridSpan w:val="2"/>') == 2, xml
    assert "gridAfter" not in xml and "gridBefore" not in xml, xml
    # beyond docx4j's assertions: which cells span 2, and every reading of the row spans the grid
    choice, fallback = branches_of_the_element(pkg)
    assert spans_in(choice) == [2] and spans_in(fallback) == [2]
    assert columns_through_each_branch(pkg) == [3, 3]


def test_rejecting_a_grid_change_counts_a_cell_after_the_element_from_its_old_span() -> None:
    """Beyond docx4j's test: the cell outside the element after it, where docx4j's shape has it before.

    Row 1 is the element first, C | F over two old columns, then X over four;
    row 2 a, b and c over two each. Rejected, the grid is 3 x 1000, X spans 2,
    the others 1. Before ``84b202c5a`` the Fallback's layout wrote X's span 2
    (new columns), and Word's then counted X from that, two old columns where
    it spans four: X got span 1, the row two columns through either branch in
    a grid of three. Here the cell is miscounted itself, where in docx4j's
    shape it is the cell after it, placed from the wrong edge.
    """
    pkg = with_body(
        regridded_in_500s(
            f'<w:tr>{element(spanned_in_500s("C", 2), spanned_in_500s("F", 2))}{spanned_in_500s("X", 4)}</w:tr>'
            f'<w:tr>{spanned_in_500s("a", 2)}{spanned_in_500s("b", 2)}{spanned_in_500s("c", 2)}</w:tr>'
        )
    )
    (change,) = revisions.tracked_changes(pkg)
    change.reject()
    xml = body_xml(pkg)
    assert grid_of(pkg) == [1000, 1000, 1000]
    assert xml.count("<w:gridSpan ") == 1 and xml.count('<w:gridSpan w:val="2"/>') == 1, xml
    choice, fallback = branches_of_the_element(pkg)
    assert spans_in(choice) == [1] and spans_in(fallback) == [1]
    assert _span(first_row(pkg).content[1]) == 2, "X, after the element"
    assert columns_through_each_branch(pkg) == [3, 3]
    assert "gridAfter" not in xml and "gridBefore" not in xml, xml
    assert revisions.tracked_changes(pkg) == []


# -- the text of a table's change: the branch Word reads (docx4j 84b202c5a) ------


def test_a_table_changes_text_reads_the_branch_word_reads() -> None:
    """docx4j ``aTableChangesTextReadsTheBranchWordReads`` (this package's finding, CR-012 section 10.11).

    A whole table's change's text is what Word shows: the rows of the branch it
    reads of a row-level element, not every branch's. ``901a922a3`` had read
    every branch's.
    """
    pkg = with_body(
        f"<w:tbl>{CHANGED_TBL_PR}"
        '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>'
        f"{element(plain_row('C'), plain_row('F'))}</w:tbl>"
    )
    changes = revisions.tracked_changes(pkg)
    assert len(changes) == 1, changes
    text = changes[0].text
    assert "C" in text, text
    assert "F" not in text, text


def test_a_whole_tables_change_text_reads_the_branch_word_reads() -> None:
    """The pin of ``901a922a3``'s reading (CR-012 section 10.11), turned into the expectation (section 10.12).

    The table's grid changed, so the change's text is every row's that Word
    shows: the Choice's, ``A\\tB\\r\\n``. ``901a922a3`` gave the Fallback's too,
    ``A\\tB\\r\\nX\\tY\\r\\n``.
    """
    pkg = with_body(
        regridded(
            4,
            f'<mc:AlternateContent><mc:Choice Requires="w14"><w:tr>{spanned("A", 1)}{spanned("B", 3)}</w:tr></mc:Choice>'
            f'<mc:Fallback><w:tr>{spanned("X", 2)}{spanned("Y", 2)}</w:tr></mc:Fallback></mc:AlternateContent>',
        )
    )
    (change,) = revisions.tracked_changes(pkg)
    assert change.text == "A\tB\r\n", "the Choice's row alone"


def test_word_mode_lists_and_resolves_a_table_property_record_as_before() -> None:
    """In ``word`` mode the element is replaced on load by Word's branch: one cell, one record, resolved as before."""
    for accept in (True, False):
        pkg = load(with_body(cell_record_in_each_branch()).save(), options=LoadOptions(mc_block_level="word"))
        (change,) = revisions.tracked_changes(pkg)
        assert change.twins == []
        assert len(pkg.main_document_part.resolved_alternate_content) == 1, "recorded when the part was read"
        if accept:
            change.accept()
        else:
            change.reject()
        xml = body_xml(pkg)
        assert "AlternateContent" not in xml and "tcPrChange" not in xml
        assert ">C<" in xml and ">F<" not in xml
        assert xml.count('w:fill="FF0000"') == (1 if accept else 0)
        assert xml.count('<w:tcW w:w="1000"') == 1
        assert revisions.tracked_changes(pkg) == []
