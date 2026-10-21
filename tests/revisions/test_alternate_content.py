"""Tracked changes inside a block- and a paragraph-level ``mc:AlternateContent``. CR-012 Phase A, section 4.1 item 6.

The port of docx4j's ``TrackedChangesAlternateContentTest`` (``2853b7280``, its
CR-051 section 10.3 item 3 and 10.5 item 3): the changes of the branch Word
reads are listed once, each with the other branch's change of its kind at its
position as a twin, which resolves with it; ``accept_all``'s listing walks every
branch and counts the read branch's. The trees are built outside a part
(``set_contents``), so the elements stay whole whatever the load mode.

Before this phase the engine read such an element as holding nothing: no change
in it was listed, accepted or rejected.
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
