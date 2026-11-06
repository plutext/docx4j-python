"""CR-038: a ``w14:paraId`` on every paragraph and row added, where the part uses them.

    .venv-fork/bin/python -m pytest tests/content/test_para_ids.py -q

Word renumbers every id in a file with no comments if one paragraph or row lacks one
(CR-038 section 8). The oracle here is that rule, so these tests hold the count, from a
Word save with ids on everything, on the saved bytes. Only Word can hold the outcome,
which is CR-038's Word check.
"""

from __future__ import annotations

import base64
import inspect
import io
import re
import sys
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any, NamedTuple

import pytest
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py import create_package, load
from docx4j_py.docx import Document
from docx4j_py.docx import _base as facade_base
from docx4j_py.docx import comments as facade_comments
from docx4j_py.docx import document as facade_document
from docx4j_py.docx import section as facade_section
from docx4j_py.docx import table as facade_table
from docx4j_py.docx import text as facade_text
from docx4j_py.model import content
from docx4j_py.model.content import addresses
from docx4j_py.wml import el

CHECK76 = ROOT / "tests" / "fixtures" / "check76"
#: Word's save, with an id on every paragraph and row in the body, the header, the notes and the comments.
WORD_SAVE = CHECK76 / "76t18-hosts-control-word365.docx"
#: Word's save with no header, footer or comment.
NO_HEADER = CHECK76 / "76t26-more-hosts-control-word365.docx"
W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_URI}}}"
W14_URI = "http://schemas.microsoft.com/office/word/2010/wordml"
ID = f"{{{W14_URI}}}paraId"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
STORY = re.compile(r"word/(document|header\d*|footer\d*|comments|footnotes|endnotes)\.xml$")
PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==")

P_XML = "<w:p><w:r><w:t>new</w:t></w:r></w:p>"
TBL_XML = (
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>'
    "<w:tr><w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"
)
SDT_XML = "<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>"
VML = 'xmlns:v="urn:schemas-microsoft-com:vml"'
TEXT_BOX_XML = f"<w:p><w:r><w:pict {VML}><v:shape><v:textbox><w:txbxContent><w:p/></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"


def saved(pkg: Any) -> bytes:
    out = io.BytesIO()
    pkg.save(out, **({} if isinstance(pkg, Document) else {"check": False}))
    return out.getvalue()


def census(data: bytes) -> dict[str, list[Any]]:
    """Each story part's paragraphs and rows, each one's id or None."""
    archive = zipfile.ZipFile(io.BytesIO(data))
    found: dict[str, list[Any]] = {}
    for name in archive.namelist():
        if STORY.match(name):
            root = etree.fromstring(archive.read(name))
            found[name] = [item.get(ID) for item in root.iter(W + "p", W + "tr")]
    return found


def gaps(data: bytes) -> dict[str, int]:
    return {name: ids.count(None) for name, ids in census(data).items() if None in ids}


def repeats(data: bytes) -> dict[str, int]:
    found = {}
    for name, ids in census(data).items():
        held = [value for value in ids if value]
        if len(held) != len(set(held)):
            found[name] = len(held) - len(set(held))
    return found


def fresh() -> Any:
    pkg = load(WORD_SAVE)
    assert pkg.main_document_part.contents is not None
    return pkg


def first_text(pkg: Any) -> Any:
    return next(paragraph for paragraph in pkg.body.paragraphs if paragraph.text.strip())


def table(pkg: Any) -> Any:
    return pkg.body.tables[0]


def cell(pkg: Any) -> Any:
    return table(pkg).rows[0].cells[0]


def control(pkg: Any) -> Any:
    return next(item for item in pkg.body.content_controls if item.form == "block")


def tracked(pkg: Any) -> Any:
    pkg.change_tracking_mode = "TrackAll"
    return pkg


class Route(NamedTuple):
    """A way of adding, what public method it is, and whether it takes something away as it goes."""

    label: str
    method: str
    run: Callable[[Any], Any]
    replaces: bool = False


ROUTES = [
    Route("insert_paragraph End", "Body.insert_paragraph", lambda k: k.body.insert_paragraph("x", location="End")),
    Route("insert_paragraph Start", "Body.insert_paragraph", lambda k: k.body.insert_paragraph("x", location="Start")),
    Route("insert_table End", "Body.insert_table", lambda k: k.body.insert_table(2, 2, location="End")),
    Route("insert_table Start", "Body.insert_table", lambda k: k.body.insert_table(2, 2, location="Start")),
    Route("insert_table values", "Body.insert_table", lambda k: k.body.insert_table(2, 2, values=[["a", "b"], ["c", "d"]])),
    Route("tracked insert_table", "Body.insert_table", lambda k: tracked(k).body.insert_table(2, 2)),
    Route("tracked insert_paragraph", "Body.insert_paragraph", lambda k: tracked(k).body.insert_paragraph("x")),
    Route("insert_break Page", "Body.insert_break", lambda k: k.body.insert_break("Page", location="End")),
    Route("insert_break SectionNext", "Body.insert_break", lambda k: k.body.insert_break("SectionNext", location="End")),
    Route("Body.insert_text", "Body.insert_text", lambda k: k.body.insert_text("a\nb", location="End")),
    Route("Body.insert_ooxml w:p", "Body.insert_ooxml", lambda k: k.body.insert_ooxml(P_XML)),
    Route("Body.insert_ooxml w:tbl", "Body.insert_ooxml", lambda k: k.body.insert_ooxml(TBL_XML)),
    Route("Body.insert_ooxml w:sdt", "Body.insert_ooxml", lambda k: k.body.insert_ooxml(SDT_XML)),
    Route("Body.insert_ooxml text box", "Body.insert_ooxml", lambda k: k.body.insert_ooxml(TEXT_BOX_XML)),
    Route("Body.insert_xml w:p", "Body.insert_xml", lambda k: k.body.insert_xml(P_XML)),
    Route("Body.insert_xml w:tbl", "Body.insert_xml", lambda k: k.body.insert_xml(TBL_XML)),
    Route("Markdown paragraphs", "Body.insert_markdown", lambda k: k.body.insert_markdown("a\n\nb")),
    Route("Markdown table", "Body.insert_markdown", lambda k: k.body.insert_markdown("|a|b|\n|-|-|\n|c|d|\n")),
    Route("Markdown list", "Body.insert_markdown", lambda k: k.body.insert_markdown("- a\n- b\n")),
    Route("Markdown code", "Body.insert_markdown", lambda k: k.body.insert_markdown("```\na\nb\n```\n")),
    Route("Markdown rule", "Body.insert_markdown", lambda k: k.body.insert_markdown("a\n\n---\n\nb\n")),
    Route("Body.insert_content_control", "Body.insert_content_control", lambda k: k.body.insert_content_control()),
    Route("Body.insert_inline_picture", "Body.insert_inline_picture", lambda k: k.body.insert_inline_picture(PNG)),
    Route(
        "Body.insert_inline_picture_from_base64",
        "Body.insert_inline_picture_from_base64",
        lambda k: k.body.insert_inline_picture_from_base64(base64.b64encode(PNG).decode()),
    ),
    Route("add_paragraph_of_text", "Body.add_paragraph_of_text", lambda k: k.body.add_paragraph_of_text("x")),
    Route("add_styled_paragraph_of_text", "Body.add_styled_paragraph_of_text", lambda k: k.body.add_styled_paragraph_of_text("Normal", "x")),
    Route("add_object", "Body.add_object", lambda k: k.body.add_object(el.p())),
    Route("insert_element", "Body.insert_element", lambda k: k.body.insert_element(el.p())),
    Route("Paragraph.insert_paragraph After", "Paragraph.insert_paragraph", lambda k: first_text(k).insert_paragraph("x", location="After")),
    Route("Paragraph.insert_paragraph Before", "Paragraph.insert_paragraph", lambda k: first_text(k).insert_paragraph("x", location="Before")),
    Route("Paragraph.insert_text", "Paragraph.insert_text", lambda k: first_text(k).insert_text("a\nb", location="End")),
    Route("Paragraph.insert_break Page", "Paragraph.insert_break", lambda k: first_text(k).insert_break("Page", location="After")),
    Route("Paragraph.insert_break SectionNext", "Paragraph.insert_break", lambda k: first_text(k).insert_break("SectionNext", location="After")),
    Route("Paragraph.insert_ooxml w:p", "Paragraph.insert_ooxml", lambda k: first_text(k).insert_ooxml(P_XML)),
    Route("Paragraph.insert_ooxml w:tbl", "Paragraph.insert_ooxml", lambda k: first_text(k).insert_ooxml(TBL_XML)),
    Route("Paragraph.insert_xml", "Paragraph.insert_xml", lambda k: first_text(k).insert_xml(P_XML)),
    Route("Paragraph.insert_markdown", "Paragraph.insert_markdown", lambda k: first_text(k).insert_markdown("a\n\nb")),
    Route("Paragraph.insert_content_control", "Paragraph.insert_content_control", lambda k: first_text(k).insert_content_control()),
    Route("Paragraph.insert_inline_picture", "Paragraph.insert_inline_picture", lambda k: first_text(k).insert_inline_picture(PNG)),
    Route(
        "Paragraph.insert_inline_picture_from_base64",
        "Paragraph.insert_inline_picture_from_base64",
        lambda k: first_text(k).insert_inline_picture_from_base64(base64.b64encode(PNG).decode()),
    ),
    Route("Paragraph.insert_comment", "Paragraph.insert_comment", lambda k: first_text(k).insert_comment("x")),
    Route("Paragraph.insert_comment two lines", "Paragraph.insert_comment", lambda k: first_text(k).insert_comment("a\nb")),
    Route("Range.insert_paragraph After", "Range.insert_paragraph", lambda k: first_text(k).get_range().insert_paragraph("x", location="After")),
    Route("Range.insert_paragraph Before", "Range.insert_paragraph", lambda k: first_text(k).get_range().insert_paragraph("x", location="Before")),
    Route("Range.insert_text Replace", "Range.insert_text", lambda k: first_text(k).get_range().insert_text("a\rb", location="Replace")),
    Route("Range.insert_ooxml w:p", "Range.insert_ooxml", lambda k: first_text(k).get_range().insert_ooxml(P_XML, location="Replace"), True),
    Route("Range.insert_ooxml w:tbl", "Range.insert_ooxml", lambda k: first_text(k).get_range().insert_ooxml(TBL_XML, location="Replace"), True),
    Route("Range.insert_comment", "Range.insert_comment", lambda k: first_text(k).get_range().insert_comment("x")),
    Route("Range.insert_content_control", "Range.insert_content_control", lambda k: first_text(k).get_range().insert_content_control()),
    Route("add_rows End", "Table.add_rows", lambda k: table(k).add_rows(2, location="End")),
    Route("add_rows Start", "Table.add_rows", lambda k: table(k).add_rows(2, location="Start")),
    Route("add_rows values", "Table.add_rows", lambda k: table(k).add_rows(1, values=[["a"] * len(table(k).rows[0].cells)])),
    Route("tracked add_rows", "Table.add_rows", lambda k: (tracked(k), table(k).add_rows(1, location="End"))),
    Route("insert_rows After", "TableRow.insert_rows", lambda k: table(k).rows[0].insert_rows(2, location="After")),
    Route("insert_rows Before", "TableRow.insert_rows", lambda k: table(k).rows[0].insert_rows(2, location="Before")),
    Route("TableCell.insert_paragraph", "TableCell.insert_paragraph", lambda k: cell(k).insert_paragraph("x")),
    Route("TableCell.insert_text", "TableCell.insert_text", lambda k: cell(k).insert_text("a\nb")),
    Route("a nested table", "Body.insert_table", lambda k: cell(k).body.insert_table(2, 1, location="End")),
    Route("a nested fragment", "Body.insert_ooxml", lambda k: cell(k).body.insert_ooxml(TBL_XML)),
    Route("a new row's paragraph", "TableCell.insert_paragraph", lambda k: table(k).add_rows(1, location="End")[0].cells[0].insert_paragraph("x")),
    Route("ContentControl.insert_paragraph", "ContentControl.insert_paragraph", lambda k: control(k).insert_paragraph("x")),
    Route("ContentControl.insert_text", "ContentControl.insert_text", lambda k: control(k).insert_text("a\nb")),
    Route("ContentControl.insert_copy_after", "ContentControl.insert_copy_after", lambda k: control(k).insert_copy_after()),
    Route("a table in a control", "Body.insert_table", lambda k: control(k).body().insert_table(2, 1, location="End")),
    Route("a new control's paragraph", "ContentControl.insert_paragraph", lambda k: k.body.insert_content_control().insert_paragraph("x")),
    Route("List.insert_paragraph", "List.insert_paragraph", lambda k: first_text(k).start_new_list().insert_paragraph("x")),
    Route("Comment.reply", "Comment.add_reply", lambda k: first_text(k).insert_comment("x").reply("y")),
    Route("Comment.content", "Comment.add_reply", lambda k: setattr(first_text(k).insert_comment("x"), "content", "a\nb")),
    Route("a paragraph copied", "Body.insert_xml", lambda k: k.body.insert_xml(first_text(k).get_xml())),
    Route("a table copied", "Body.insert_xml", lambda k: k.body.insert_xml(table(k).get_xml())),
]
FACADE_ROUTES = [
    Route("add_paragraph", "Document.add_paragraph", lambda d: d.add_paragraph("x")),
    Route("add_heading", "Document.add_heading", lambda d: d.add_heading("x", 1)),
    Route("add_page_break", "Document.add_page_break", lambda d: d.add_page_break()),
    Route("add_section", "Document.add_section", lambda d: d.add_section()),
    Route("add_picture", "Document.add_picture", lambda d: d.add_picture(io.BytesIO(PNG))),
    Route("add_table", "Document.add_table", lambda d: d.add_table(2, 2)),
    Route("add_comment", "Document.add_comment", lambda d: d.add_comment(d.paragraphs[1].runs[0], "x")),
    Route("add_row", "Table.add_row", lambda d: d.tables[0].add_row()),
    Route("add_column", "Table.add_column", lambda d: d.tables[0].add_column(100)),
    Route("cell add_paragraph", "_Cell.add_paragraph", lambda d: d.tables[0].cell(0, 0).add_paragraph("x")),
    Route("cell add_table", "_Cell.add_table", lambda d: d.tables[0].cell(0, 0).add_table(2, 1)),
    Route("insert_paragraph_before", "Paragraph.insert_paragraph_before", lambda d: d.paragraphs[1].insert_paragraph_before("x")),
    Route("header add_paragraph", "_Header.add_paragraph", lambda d: d.sections[0].header.add_paragraph("x")),
    Route("header add_table", "_Header.add_table", lambda d: d.sections[0].header.add_table(1, 1, 100)),
    Route("a new footer", "_Footer.add_paragraph", lambda d: d.sections[0].footer.add_paragraph("x")),
    Route("a table in a new footer", "_Footer.add_table", lambda d: d.sections[0].footer.add_table(1, 1, 100)),
    Route("comment add_paragraph", "Comment.add_paragraph", lambda d: d.add_comment(d.paragraphs[1].runs[0], "x").add_paragraph("y")),
    Route("comment add_table", "Comment.add_table", lambda d: d.add_comment(d.paragraphs[1].runs[0], "x").add_table(1, 1)),
]
#: Public methods that add nothing of the two, each with why.
ADDS_NEITHER = {
    "Paragraph.insert_items_at": "items of a paragraph",
    "Paragraph.add_picture": "a run",
    "Paragraph.add_run": "a run",
    "Run.add_break": "in a run",
    "Run.add_picture": "in a run",
    "Run.add_tab": "in a run",
    "Run.add_text": "in a run",
    "Comments.add_comment": "as Document.add_comment",
    "BlockItemContainer.add_paragraph": "the base of the routes listed",
    "BlockItemContainer.add_table": "the base of the routes listed",
}
CLASSES = {
    **{name: getattr(content, name) for name in ("Body", "Paragraph", "Range", "Table", "TableRow", "TableCell", "ContentControl", "List", "Comment")},
}
FACADE_CLASSES = {
    "Document": facade_document.Document,
    "BlockItemContainer": facade_base.BlockItemContainer,
    "Table": facade_table.Table,
    "_Cell": facade_table._Cell,
    "Paragraph": facade_text.Paragraph,
    "Run": facade_text.Run,
    "_Header": facade_section._Header,
    "_Footer": facade_section._Footer,
    "Comment": facade_comments.Comment,
    "Comments": facade_comments.Comments,
}


def public_routes(classes: dict[str, type]) -> set[str]:
    return {
        f"{name}.{method}"
        for name, clazz in classes.items()
        for method, _member in inspect.getmembers(clazz, inspect.isfunction)
        if re.match(r"(insert_|add_|append_)", method)
    }


def test_every_public_route_is_listed() -> None:
    """A new route fails this until it's in a table here: the gap came from routes nobody listed."""
    listed = {route.method for route in ROUTES} | {route.method for route in FACADE_ROUTES} | set(ADDS_NEITHER)
    assert public_routes(CLASSES) | public_routes(FACADE_CLASSES) <= listed


def check_route(route: Route, data: bytes, before: dict[str, list[Any]]) -> None:
    assert gaps(data) == {}, route.label
    assert repeats(data) == {}, route.label
    after = {value for ids in census(data).values() for value in ids}
    lost = {value for ids in before.values() for value in ids} - after
    assert len(lost) <= (1 if route.replaces else 0), route.label


@pytest.mark.parametrize("route", ROUTES, ids=lambda route: route.label)
def test_a_route_leaves_each_paragraph_and_row_one_id(route: Route) -> None:
    before = census(WORD_SAVE.read_bytes())
    pkg = fresh()
    route.run(pkg)
    check_route(route, saved(pkg), before)


@pytest.mark.parametrize("route", FACADE_ROUTES, ids=lambda route: route.label)
def test_a_facade_route_leaves_each_paragraph_and_row_one_id(route: Route) -> None:
    before = census(WORD_SAVE.read_bytes())
    document = Document(WORD_SAVE)
    route.run(document)
    check_route(route, saved(document), before)


def test_reading_and_saving_changes_no_id() -> None:
    assert census(saved(fresh())) == census(WORD_SAVE.read_bytes())


# ---------------------------------------------------------------------------
# the story decides
# ---------------------------------------------------------------------------


def test_a_cell_with_no_ids_of_its_own_reads_as_the_document() -> None:
    """A cell's ``Body`` sees only the cell: the part decides."""
    pkg = fresh()
    new_cell = table(pkg).add_rows(1, location="End")[0].cells[0]
    for item in addresses.identified(new_cell.element):
        item.para_id = None
    pkg._para_ids_taken = None  # as a fresh read would find it, a part with ids and this cell with none
    assert new_cell.insert_paragraph("x").para_id is not None


def test_taken_holds_rows_and_text_boxes() -> None:
    pkg = fresh()
    row = table(pkg).rows[0].tr.para_id
    assert row.upper() in addresses.para_ids_of(pkg.body)
    assert row.upper() in addresses.para_ids_of(cell(pkg).body)  # the part's, from anywhere in it


def test_the_set_is_built_once_a_part(monkeypatch: pytest.MonkeyPatch) -> None:
    """Two thousand new paragraphs stay linear."""
    calls = []
    real = addresses.para_ids_of
    monkeypatch.setattr(addresses, "para_ids_of", lambda body: calls.append(1) or real(body))
    pkg = fresh()
    for _ in range(40):
        pkg.body.insert_paragraph("x")
    table(pkg).add_rows(3, location="End")
    cell(pkg).insert_paragraph("x")  # a body of its own, in the same part
    control(pkg).insert_paragraph("x")
    assert len(calls) == 1


def test_a_clean_insert_never_walks_the_part_for_collisions(monkeypatch: pytest.MonkeyPatch) -> None:
    """Only an id already in the set costs a walk, so a fragment with none pays nothing."""
    pkg = fresh()
    pkg.body.insert_paragraph("x")  # the set is built
    real = addresses._story_of
    calls = []
    monkeypatch.setattr(addresses, "_story_of", lambda body: calls.append(1) or real(body))
    pkg.body.insert_ooxml(TBL_XML)
    assert calls == []


# ---------------------------------------------------------------------------
# an id already there (Q1)
# ---------------------------------------------------------------------------

W14 = f'xmlns:w14="{W14_URI}"'


def with_id(para_id: str, text: str = "new") -> str:
    return f'<w:p {W14} w14:paraId="{para_id}"><w:r><w:t>{text}</w:t></w:r></w:p>'


def test_a_fragments_own_free_id_stays() -> None:
    pkg = fresh()
    pkg.body.insert_ooxml(with_id("1A2B3C4D"))
    assert "1A2B3C4D" in census(saved(pkg))["word/document.xml"]


def test_an_id_the_part_holds_is_replaced_and_the_original_keeps_it() -> None:
    pkg = fresh()
    original = first_text(pkg)
    held = original.para_id
    (added,) = pkg.body.insert_ooxml(with_id(held))
    assert added.para_id != held and original.para_id == held
    assert pkg.paragraph_at(para_id=held).text == original.text


def test_the_copy_is_the_one_replaced_wherever_it_lands() -> None:
    """Word gives the later of two a new id. Here the copy goes, even before its original, for the address a caller holds."""
    pkg = fresh()
    original = first_text(pkg)
    held = original.para_id
    (added,) = pkg.body.insert_ooxml(original.get_xml(), location="Start")
    assert original.para_id == held and added.para_id != held


def test_a_repeat_inside_a_fragment_is_replaced() -> None:
    pkg = fresh()
    first, second = pkg.body.insert_ooxml(with_id("1A2B3C4D", "a") + with_id("1A2B3C4D", "b"))
    assert first.para_id == "1A2B3C4D" and second.para_id not in (None, "1A2B3C4D")


def test_a_free_id_is_noted_for_the_next_insert() -> None:
    pkg = fresh()
    pkg.body.insert_paragraph("x")  # the set is built, so what comes next isn't in it
    (first,) = pkg.body.insert_ooxml(with_id("1A2B3C4D", "a"))
    (second,) = pkg.body.insert_ooxml(with_id("1A2B3C4D", "b"))
    assert first.para_id == "1A2B3C4D" and second.para_id != "1A2B3C4D"


def test_three_of_one_id_in_one_block_leave_it_on_one() -> None:
    """The walker works down what it found, so the id isn't lost altogether."""
    cells = "".join(f'<w:tc><w:p {W14} w14:paraId="1A2B3C4D"/></w:tc>' for _ in range(3))
    pkg = fresh()
    pkg.body.insert_ooxml(f'<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1"/></w:tblGrid><w:tr>{cells}</w:tr></w:tbl>')
    data = saved(pkg)
    assert census(data)["word/document.xml"].count("1A2B3C4D") == 1 and repeats(data) == {}


def test_a_paragraph_replaced_by_one_carrying_its_id_keeps_it() -> None:
    """The set isn't reduced on a delete, so the tree decides: nothing else holds the id by then."""
    pkg = fresh()
    paragraph = first_text(pkg)
    held = paragraph.para_id
    xml = paragraph.get_xml()
    paragraph.delete()
    pkg.body.insert_ooxml(xml)
    assert held in census(saved(pkg))["word/document.xml"]
    assert repeats(saved(pkg)) == {}


def test_a_controls_first_paragraph_made_for_it_holds_an_id() -> None:
    pkg = fresh()
    pkg.body.insert_ooxml('<w:sdt><w:sdtPr><w:tag w:val="empty"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent/></w:sdt>', location="Start")
    empty = next(item for item in pkg.body.content_controls if item.tag == "empty")
    assert not list(addresses.identified(empty.element))  # it holds none, so the write has to make one
    empty.insert_text("x")
    assert gaps(saved(pkg)) == {}


def test_a_placeholders_further_paragraph_holds_an_id() -> None:
    """A block control shows an entry's second paragraph too."""
    from test_placeholders_words_form import _with_entry_of_two_paragraphs

    pkg, name = _with_entry_of_two_paragraphs()
    pkg.assigns_para_ids = True
    block = next(item for item in pkg.body.content_controls if item.form == "block" and item.placeholder_entry == name)
    block.show_placeholder()
    shown = list(addresses.identified(block.element))  # paragraphs hides what a placeholder shows
    assert len(shown) == 2 and shown[1].para_id is not None


def test_a_control_copy_takes_new_ids() -> None:
    pkg = fresh()
    before = {item.para_id for item in addresses.identified(control(pkg).element)}
    copy = control(pkg).insert_copy_after()
    assert before.isdisjoint({item.para_id for item in addresses.identified(copy.element)})


ALTERNATE = (
    '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    f'<mc:Choice xmlns:w14="{W14_URI}" Requires="w14"><w:pict {VML}><v:shape><v:textbox><w:txbxContent>{{choice}}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Choice>'
    f"<mc:Fallback><w:pict {VML}><v:shape><v:textbox><w:txbxContent>{{fallback}}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>"
    "</mc:AlternateContent></w:r></w:p>"
)


def branch_ids(data: bytes) -> tuple[list[Any], list[Any]]:
    root = etree.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml"))
    choice = [p.get(ID) for c in root.iter(MC + "Choice") for p in c.iter(W + "p")]
    fallback = [p.get(ID) for c in root.iter(MC + "Fallback") for p in c.iter(W + "p")]
    return choice, fallback


def test_a_fallback_isnt_entered() -> None:
    """Word reads the Choice's id and writes it over the Fallback's (CR-038 section 8)."""
    pkg = fresh()
    pkg.body.insert_ooxml(ALTERNATE.format(choice="<w:p/>", fallback="<w:p/>"))
    choice, fallback = branch_ids(saved(pkg))
    assert choice[0] is not None and fallback == [None]


def test_a_pair_as_word_writes_it_is_left() -> None:
    pkg = fresh()
    pair = f'<w:p {W14} w14:paraId="1A2B3C4D"/>'
    pkg.body.insert_ooxml(ALTERNATE.format(choice=pair, fallback=pair))
    assert branch_ids(saved(pkg)) == (["1A2B3C4D"], ["1A2B3C4D"])


# ---------------------------------------------------------------------------
# other parts (Q2)
# ---------------------------------------------------------------------------


def test_a_new_footer_holds_ids_where_the_main_part_does() -> None:
    document = Document(WORD_SAVE)
    document.sections[0].footer.add_paragraph("x")
    assert census(saved(document))["word/footer1.xml"].count(None) == 0


def test_a_new_header_holds_ids_where_the_main_part_does() -> None:
    document = Document(NO_HEADER)
    document.sections[0].header.add_paragraph("x")
    assert census(saved(document))["word/header1.xml"].count(None) == 0


def no_ids() -> Any:
    pkg = create_package()
    pkg.assigns_para_ids = None
    pkg.main_document_part.set_bytes(
        f'<w:document xmlns:w="{W_URI}"><w:body><w:p><w:r><w:t>a</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'.encode()
    )
    assert pkg.main_document_part.contents is not None
    return pkg


@pytest.mark.parametrize(
    "run",
    [
        lambda k: k.body.insert_paragraph("x"),
        lambda k: k.body.insert_table(2, 2),
        lambda k: k.body.insert_ooxml(TBL_XML),
        lambda k: k.body.insert_markdown("|a|b|\n|-|-|\n|c|d|\n"),
        lambda k: k.body.tables[0].add_rows(1) if k.body.tables else k.body.insert_table(1, 1).add_rows(1),
    ],
)
def test_a_document_using_no_ids_stays_so(run: Callable[[Any], Any]) -> None:
    pkg = no_ids()
    run(pkg)
    assert not any(value for value in census(saved(pkg))["word/document.xml"])


def test_a_new_footer_in_a_document_with_no_ids_takes_none() -> None:
    pkg = no_ids()
    document = Document(io.BytesIO(saved(pkg)))
    document.sections[0].footer.add_paragraph("x")
    assert not any(value for value in census(saved(document))["word/footer1.xml"])


def test_an_existing_part_with_no_ids_stays_so() -> None:
    """It's the part that decides. Only a header or footer made here asks the main part."""
    out = io.BytesIO()
    with zipfile.ZipFile(WORD_SAVE) as source, zipfile.ZipFile(out, "w") as target:
        for info in source.infolist():
            data = source.read(info)
            if info.filename == "word/header1.xml":
                data = re.sub(rb' w14:(paraId|textId)="[^"]*"', b"", data)
            target.writestr(info, data)
    document = Document(io.BytesIO(out.getvalue()))
    document.sections[0].header.add_paragraph("x")
    assert not any(census(saved(document))["word/header1.xml"])


def test_a_new_part_asks_without_marking_the_main_part() -> None:
    pkg = load(WORD_SAVE)
    assert addresses._main_uses_ids(pkg)
    assert pkg.main_document_part._contents is None


def test_the_glossary_takes_no_ids() -> None:
    """Word writes none there, and drops one it's given (CR-038 section 8)."""
    pkg = fresh()
    control(pkg).show_placeholder()
    root = etree.fromstring(zipfile.ZipFile(io.BytesIO(saved(pkg))).read("word/glossary/document.xml"))
    assert [p.get(ID) for p in root.iter(W + "p")] == [None]


@pytest.mark.parametrize("write", [lambda k: first_text(k).insert_comment("a\nb\nc"), lambda k: setattr(first_text(k).insert_comment("x"), "content", "a\nb\nc")])
def test_every_paragraph_of_a_comment_holds_an_id(write: Callable[[Any], Any]) -> None:
    pkg = fresh()
    write(pkg)
    data = saved(pkg)
    assert gaps(data) == {} and repeats(data) == {}
    assert len(census(data)["word/comments.xml"]) >= 3


def test_a_comment_with_no_id_of_its_own_stays_so() -> None:
    pkg = fresh()
    comment = first_text(pkg).insert_comment("x")
    comment.element.content[0].para_id = None
    comment.content = "a\nb"
    assert [item.para_id for item in comment.element.content] == [None, None]


# ---------------------------------------------------------------------------
# created documents, ensure_para_ids, the report, determinism
# ---------------------------------------------------------------------------


def test_a_created_document_holds_ids_on_everything() -> None:
    pkg = create_package()
    pkg.body.insert_paragraph("x")
    pkg.body.insert_table(2, 2).add_rows(1)
    assert gaps(saved(pkg)) == {}


def test_a_created_documents_new_footer_holds_ids() -> None:
    document = Document()
    document.sections[0].footer.add_paragraph("x")
    assert gaps(saved(document)) == {}


def test_ensure_para_ids_reaches_rows_and_text_boxes() -> None:
    pkg = no_ids()
    pkg.body.insert_table(2, 1)
    pkg.body.insert_ooxml(TEXT_BOX_XML)
    expected = sum(len(ids) for ids in census(saved(pkg)).values())
    assigned = pkg.body.ensure_para_ids()
    data = saved(pkg)
    assert gaps(data) == {} and len(assigned) == expected


def test_ensure_para_ids_is_in_document_order() -> None:
    pkg = no_ids()
    pkg.body.insert_table(1, 1)
    assigned = pkg.body.ensure_para_ids()
    root = etree.fromstring(zipfile.ZipFile(io.BytesIO(saved(pkg))).read("word/document.xml"))
    assert assigned == [item.get(ID) for item in root.iter(W + "p", W + "tr")]


def test_ensure_para_ids_leaves_a_documents_own_repeats() -> None:
    pkg = fresh()
    pkg.body.paragraphs[1].para_id = pkg.body.paragraphs[0].para_id
    pkg.body.ensure_para_ids()
    assert pkg.body.paragraphs[1].para_id == pkg.body.paragraphs[0].para_id


def test_the_report_holds_every_id_a_call_gave() -> None:
    pkg = fresh()
    pkg.body.insert_table(2, 2)
    assert len(pkg.last_change.created_para_ids) == 6  # two rows, and four paragraphs


def test_the_same_seed_and_calls_give_the_same_ids() -> None:
    def ids() -> dict[str, list[Any]]:
        pkg = fresh()
        pkg.id_seed = 7
        pkg.body.insert_table(2, 2)
        table(pkg).add_rows(1, location="End")
        return census(saved(pkg))

    assert ids() == ids()


def test_an_id_is_in_words_range() -> None:
    pkg = fresh()
    pkg.body.insert_table(3, 3)
    assert all(0 < int(value, 16) < 0x80000000 for value in pkg.last_change.created_para_ids)
