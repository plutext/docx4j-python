"""CR-014 section 3: a bound control Word reads as plain text holds one paragraph.

A **value control** is a block-level control with a ``w:dataBinding`` and no
type, ``w:text`` or ``w:richText``, not a container, its node not flat OPC
(section 3.1; core-ts's H1, check 56 at ``38dcae9``). Word can't open a file
where one holds a second block, and its repair loses it. So:

* every write that would add a block is refused before anything changes, tracked
  or not, but ``insert_paragraph`` into a single-line one, whose text joins the
  paragraph (3.2, Q3; H2);
* ``set_mapping`` and ``set_mapping_by_node`` refuse a control of a value kind
  already holding more (3.3, Q4);
* ``apply_bindings`` writes one paragraph over one holding more (3.4, Q5);
* one with no type or ``w:richText`` reads ``"PlainText"``, single-line (3.5a, Q9;
  H4), and a ``w:richText`` one is bound as text both ways (Q9).

The scope is check 56's, cited by commit; no check of this package's own comes
first (section 5). ``SWEEP`` is every block writer the audit found (section 9.3's
table): each is asked on each kind, tracked and not, and each must leave the tree,
the parts and the reports as they were.
"""

from __future__ import annotations

import base64
import re
from xml.sax.saxutils import escape

import pytest
from conftest import reloaded, sample

from docx4j_py import create_package, load
from docx4j_py.docx._base import Story
from docx4j_py.docx.text import Paragraph as FacadeParagraph
from docx4j_py.model.content import BindingError
from docx4j_py.model.content.controls import (
    holds_one_paragraph,
    is_bound_block_control,
    plain_text_lines_of,
    value_control_of,
)
from docx4j_py.model.customxml.bindings import is_container
from docx4j_py.traversal import element_name
from docx4j_py.wml import el
from docx4j_py.wml.sdt import W15_NS, sdt_kind_of, sdt_property

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DATA = "urn:cr014"
MAPPINGS = f"xmlns:ns0='{DATA}'"
XPATH = "/ns0:data[1]/ns0:value[1]"

#: Check 56's kinds (core-ts ``test/fixtures/check56/build.mjs``): the rule's four.
KINDS = {
    "no type": "",
    "w:text": "<w:text/>",
    "w:text multi-line": '<w:text w:multiLine="1"/>',
    "w:richText": "<w:richText/>",
}
SINGLE = [name for name in KINDS if name != "w:text multi-line"]

TWO_PARAGRAPHS = f'<w:p xmlns:w="{W}"><w:r><w:t>A</w:t></w:r></w:p><w:p xmlns:w="{W}"><w:r><w:t>B</w:t></w:r></w:p>'
ONE_PARAGRAPH = f'<w:p xmlns:w="{W}"><w:r><w:t>X</w:t></w:r></w:p>'
A_TABLE = (
    f'<w:tbl xmlns:w="{W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>'
    "<w:tr><w:tc><w:p><w:r><w:t>T</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
)
#: One pixel, for the writers that add an image part first.
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)
#: A flat OPC package of two paragraphs and a picture: what ``insert_ooxml`` takes, and it copies the image part first.
FLAT_OPC_TWO = f"""<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
</pkg:xmlData></pkg:part>
<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>
</pkg:xmlData></pkg:part>
<pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png" pkg:compression="store"><pkg:binaryData>{base64.b64encode(PNG).decode()}</pkg:binaryData></pkg:part>
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>
<w:document xmlns:w="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r><w:r><w:object r:id="rId9"/></w:r></w:p>
</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>"""


def control_xml(kind: str, *, content: str = "<w:p><w:r><w:t>Note</w:t></w:r></w:p>", extra: str = "") -> str:
    """A block control of `kind`, holding `content`."""
    return f'<w:sdt xmlns:w="{W}"><w:sdtPr><w:id w:val="9"/>{extra}{kind}</w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>'


def value_control(kind: str = "", *, value: str = "Note", tracked: bool = False):
    """A package holding one value control of `kind`, bound to a node holding `value`, and the control."""
    package = create_package()
    package.body.insert_paragraph("before")
    package.body.insert_xml(control_xml(kind, content=f"<w:p><w:r><w:t>{value}</w:t></w:r></w:p>"))
    package.body.insert_paragraph("after")
    control = package.body.content_controls[0]
    node = escape(value)
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>{node}</value></data>')
    assert control.xml_mapping.set_mapping(XPATH, MAPPINGS, part)
    if tracked:
        package.change_tracking_mode = "TrackAll"
    return package, control


def holding_two(kind: str = "", *, value: str = "bound value", binding: bool = True):
    """File 3a's shape, a file made elsewhere: a control of `kind` holding two paragraphs, bound directly."""
    package = create_package()
    package.body.insert_xml(
        control_xml(
            kind,
            content='<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>first</w:t></w:r></w:p>'
            "<w:p><w:r><w:t>second</w:t></w:r></w:p>",
        )
    )
    control = package.body.content_controls[0]
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>{escape(value)}</value></data>')
    if binding:
        control.put_property(el.dataBinding(xpath=XPATH, prefix_mappings=MAPPINGS, store_item_id=part.id))
    return package, control


def snapshot(package) -> tuple:
    """What a refusal must leave as it was: the document, every part's name, the data, the styles, the reports."""
    return (
        package.main_document_part.get_xml(),
        sorted(str(part.part_name) for part in package.parts.values()) if hasattr(package, "parts") else None,
        [part.get_xml() for part in package.custom_xml_parts.items],
        package.main_document_part.style_definitions_part.get_xml(),
        len(package.changes),
    )


def blocks_of(control) -> list[str]:
    return [(element_name(item) or "").rsplit("}", 1)[-1] for item in control.content]


def paragraph_of(control):
    return control.paragraphs[0]


def facade_paragraph(control):
    return FacadeParagraph(paragraph_of(control).element, Story(control.parent_body))


# ---------------------------------------------------------------------------
# the sweep: every block writer of the content API (section 3.2 and the audit)
# ---------------------------------------------------------------------------

#: name: (the write, Word's own error where check 56 measured one).
SWEEP = {
    # ContentControl: H2's refusal is its own test
    # Paragraph
    "Paragraph.insert_paragraph Before": (lambda c: paragraph_of(c).insert_paragraph("X", location="Before"), None),
    "Paragraph.insert_paragraph After": (lambda c: paragraph_of(c).insert_paragraph("X", location="After"), None),
    "Paragraph.insert_break Before": (lambda c: paragraph_of(c).insert_break("Page", location="Before"), None),
    "Paragraph.insert_break After": (lambda c: paragraph_of(c).insert_break("Line", location="After"), None),
    "Paragraph.insert_xml After": (lambda c: paragraph_of(c).insert_xml(ONE_PARAGRAPH), "GeneralException"),
    "Paragraph.insert_xml Before": (lambda c: paragraph_of(c).insert_xml(ONE_PARAGRAPH, location="Before"), "GeneralException"),
    "Paragraph.insert_xml End, two": (lambda c: paragraph_of(c).insert_xml(TWO_PARAGRAPHS, location="End"), "GeneralException"),
    "Paragraph.insert_xml Replace, two": (lambda c: paragraph_of(c).insert_xml(TWO_PARAGRAPHS, location="Replace"), "GeneralException"),
    "Paragraph.insert_xml a table": (lambda c: paragraph_of(c).insert_xml(A_TABLE, location="Replace"), "GeneralException"),
    "Paragraph.insert_ooxml After": (lambda c: paragraph_of(c).insert_ooxml(ONE_PARAGRAPH), "GeneralException"),
    "Paragraph.insert_ooxml End, two": (lambda c: paragraph_of(c).insert_ooxml(TWO_PARAGRAPHS, location="End"), "GeneralException"),
    "Paragraph.insert_ooxml Replace, two": (lambda c: paragraph_of(c).insert_ooxml(TWO_PARAGRAPHS, location="Replace"), "GeneralException"),
    "Paragraph.insert_ooxml a package": (lambda c: paragraph_of(c).insert_ooxml(FLAT_OPC_TWO, location="After"), "GeneralException"),
    "Paragraph.insert_markdown After": (lambda c: paragraph_of(c).insert_markdown("A"), None),
    "Paragraph.insert_markdown Before": (lambda c: paragraph_of(c).insert_markdown("A", location="Before"), None),
    "Paragraph.insert_markdown End, two": (lambda c: paragraph_of(c).insert_markdown("A\n\nB", location="End"), None),
    "Paragraph.insert_markdown a heading and a list": (lambda c: paragraph_of(c).insert_markdown("# H\n\n- a\n- b", location="Start"), None),
    "Paragraph.insert_content_control": (lambda c: paragraph_of(c).insert_content_control(), None),
    # Range
    "Range.insert_paragraph After": (lambda c: paragraph_of(c).get_range().insert_paragraph("X", location="After"), None),
    "Range.insert_paragraph Before": (lambda c: paragraph_of(c).get_range().insert_paragraph("X", location="Before"), None),
    "Range.insert_paragraph, a split": (lambda c: paragraph_of(c).search("ot")[0].insert_paragraph("X", location="Before"), None),
    "Range.insert_ooxml Replace, two": (lambda c: paragraph_of(c).search("ot")[0].insert_ooxml(TWO_PARAGRAPHS), "GeneralException"),
    "Range.insert_ooxml After, two": (lambda c: paragraph_of(c).get_range("Content").insert_ooxml(TWO_PARAGRAPHS, location="After"), "GeneralException"),
    "Range.insert_ooxml Before, a package": (lambda c: paragraph_of(c).get_range("Content").insert_ooxml(FLAT_OPC_TWO, location="Before"), "GeneralException"),
    "Range.insert_ooxml After, a table": (lambda c: c.body().get_range("Whole").insert_ooxml(A_TABLE, location="After"), "GeneralException"),
    # Body, over the control
    "Body.insert_paragraph End": (lambda c: c.body().insert_paragraph("X"), None),
    "Body.insert_paragraph Start": (lambda c: c.body().insert_paragraph("X", location="Start"), None),
    "Body.add_paragraph_of_text": (lambda c: c.body().add_paragraph_of_text("X"), None),
    "Body.add_styled_paragraph_of_text": (lambda c: c.body().add_styled_paragraph_of_text("Heading1", "X"), None),
    "Body.add_object": (lambda c: c.body().add_object(el.p()), None),
    "Body.insert_element": (lambda c: c.body().insert_element(el.p()), None),
    "Body.insert_element Before": (lambda c: c.body().insert_element(el.p(), location="Before", target=paragraph_of(c)), None),
    "Body.insert_table": (lambda c: c.body().insert_table(1, 1, values=[["X"]]), "InvalidArgument"),
    "Body.insert_table Start, a style": (lambda c: c.body().insert_table(1, 1, location="Start", style="Table Grid"), "InvalidArgument"),
    "Body.insert_xml": (lambda c: c.body().insert_xml(ONE_PARAGRAPH), "GeneralException"),
    "Body.insert_ooxml": (lambda c: c.body().insert_ooxml(ONE_PARAGRAPH), "GeneralException"),
    "Body.insert_ooxml a package": (lambda c: c.body().insert_ooxml(FLAT_OPC_TWO), "GeneralException"),
    "Body.insert_ooxml Replace, two": (lambda c: c.body().insert_ooxml(TWO_PARAGRAPHS, location="Replace"), "GeneralException"),
    "Body.insert_ooxml Replace, a table": (lambda c: c.body().insert_ooxml(A_TABLE, location="Replace"), "GeneralException"),
    "Body.insert_markdown": (lambda c: c.body().insert_markdown("A\n\nB"), None),
    "Body.insert_markdown a heading": (lambda c: c.body().insert_markdown("# H"), None),
    "Body.insert_break": (lambda c: c.body().insert_break("Page"), None),
    "Body.insert_inline_picture": (lambda c: c.body().insert_inline_picture(PNG), None),
    "Body.insert_content_control": (lambda c: c.body().insert_content_control(), None),
    # the python-docx facade
    "docx Paragraph.insert_paragraph_before": (lambda c: facade_paragraph(c).insert_paragraph_before("X"), None),
}


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("kind", KINDS, ids=list(KINDS))
@pytest.mark.parametrize("writer", SWEEP, ids=list(SWEEP))
def test_every_block_writer_is_refused_before_anything_changes(writer, kind, tracked):
    package, control = value_control(KINDS[kind], tracked=tracked)
    before = snapshot(package)
    write, word_error = SWEEP[writer]
    with pytest.raises(BindingError) as refused:
        write(control)
    assert refused.value.code == "binding.one_paragraph"
    assert "holds one paragraph: Word reads it as plain text" in str(refused.value)
    assert "xml_mapping.delete()" in (refused.value.hint or "")
    if word_error is not None:
        assert word_error in str(refused.value), "Q6: Word's own error, where check 56 measured one"
    else:
        assert "core-ts check 56" not in str(refused.value)
    assert snapshot(package) == before
    assert blocks_of(control) == ["p"]


def test_the_snapshot_sees_what_these_writers_change():
    """On an unbound control the same writes go in, so the sweep's unchanged snapshot is the refusal's doing."""
    package = create_package()
    package.body.insert_xml(control_xml(""))
    control = package.body.content_controls[0]
    package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>v</value></data>')
    for write in (
        lambda: control.body().insert_inline_picture(PNG),
        lambda: control.body().insert_ooxml(FLAT_OPC_TWO),
        lambda: control.body().insert_table(1, 1, location="Start", style="Table Grid"),
        lambda: control.body().insert_markdown("# H"),
    ):
        before = snapshot(package)
        write()
        after = snapshot(package)
        assert after[0] != before[0], "the document"
        assert after[4] == before[4] + 1, "a report"
    assert snapshot(package)[1] != [], "the parts"


def test_a_writer_that_adds_a_part_first_is_refused_before_it():
    package, control = value_control("")
    names = sorted(str(part.part_name) for part in package.parts.values())
    with pytest.raises(BindingError):
        control.body().insert_inline_picture(PNG)
    with pytest.raises(BindingError):
        control.body().insert_ooxml(FLAT_OPC_TWO)
    assert sorted(str(part.part_name) for part in package.parts.values()) == names
    assert not any("media" in name for name in names)


def test_a_list_item_in_a_value_control_is_refused_a_sibling():
    package, control = value_control("")
    paragraph_of(control).start_new_list()
    before = snapshot(package)
    with pytest.raises(BindingError):
        paragraph_of(control).list.insert_paragraph("X")
    assert snapshot(package) == before


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_a_refusal_in_a_dry_run_changes_neither(tracked):
    package, _control = value_control("", tracked=tracked)
    before = snapshot(package)
    with package.dry_run() as trial:
        trial_control = trial.body.content_controls[0]
        with pytest.raises(BindingError):
            paragraph_of(trial_control).insert_paragraph("x")
        assert blocks_of(trial_control) == ["p"]
    assert snapshot(package) == before


def test_a_refusal_writes_no_report():
    package, control = value_control("")
    reports = len(package.changes)
    with pytest.raises(BindingError):
        control.body().insert_table(1, 1)
    assert len(package.changes) == reports


# ---------------------------------------------------------------------------
# insert_paragraph into a value control (H2; Q3)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("kind", SINGLE, ids=SINGLE)
@pytest.mark.parametrize(("location", "text"), [("End", "start1NEWPARA1"), ("Start", "NEWPARA1start1")])
def test_insert_paragraph_into_a_single_line_one_joins_its_paragraph(kind, tracked, location, text):
    """Check 56 cases 1, 9 and 13 (``start1NEWPARA1``); the node follows the control, core-ts's F3, where Word's took ``\\r\\n``."""
    package, control = value_control(KINDS[kind], value="start1", tracked=tracked)
    out = control.insert_paragraph("NEWPARA1", location=location)
    assert out == paragraph_of(control)
    assert out.text == text
    assert blocks_of(control) == ["p"]
    assert control.xml_mapping.custom_xml_node.text == text, "the node follows the control"
    again = reloaded(package)
    assert [p.text for p in again.body.content_controls[0].paragraphs] == [text]


def test_insert_paragraph_into_a_single_line_one_drops_a_line_break():
    _package, control = value_control("")
    assert control.insert_paragraph("a\nb").text == "Noteab"


def test_insert_paragraph_takes_a_style_on_the_paragraph_it_joined():
    _package, control = value_control("")
    out = control.insert_paragraph("x", style="Heading 1")
    assert out.style_id == "Heading1"


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("location", ["Start", "End"])
def test_insert_paragraph_into_a_multi_line_one_is_refused(tracked, location):
    """Word wrote a second paragraph there and couldn't open its own save (check 56 case 5)."""
    package, control = value_control(KINDS["w:text multi-line"], tracked=tracked)
    before = snapshot(package)
    with pytest.raises(BindingError) as refused:
        control.insert_paragraph("x", location=location)
    assert refused.value.code == "binding.one_paragraph"
    assert snapshot(package) == before


@pytest.mark.parametrize("kind", KINDS, ids=list(KINDS))
@pytest.mark.parametrize("location", ["Before", "After"])
def test_insert_paragraph_beside_a_value_control_goes_outside_it(kind, location):
    _package, control = value_control(KINDS[kind])
    out = control.insert_paragraph("beside", location=location)
    assert blocks_of(control) == ["p"]
    assert out.text == "beside"


def showing(kind: str):
    package, control = value_control(kind, value="x")
    control.xml_mapping.custom_xml_node.text = ""
    package.custom_xml_parts.apply_bindings()
    assert control.is_showing_placeholder
    return package, control


def test_insert_paragraph_into_a_single_line_one_showing_its_placeholder_takes_its_place():
    """CR-019 section 4.2's write, in one paragraph: the flag cleared, the node written."""
    _package, control = showing("")
    out = control.insert_paragraph("typed")
    assert not control.is_showing_placeholder
    assert out.text == "typed"
    assert blocks_of(control) == ["p"]
    assert control.xml_mapping.custom_xml_node.text == "typed"


def test_a_multi_line_one_showing_its_placeholder_refuses_before_the_placeholder_goes():
    package, control = showing(KINDS["w:text multi-line"])
    before = snapshot(package)
    with pytest.raises(BindingError):
        control.insert_paragraph("typed")
    assert control.is_showing_placeholder
    assert snapshot(package) == before


def test_beside_the_only_paragraph_showing_its_placeholder_stays_outside_the_control():
    """CR-019 D6 (b) keeps the rule: the new paragraph is not in the control."""
    _package, control = showing("")
    out = paragraph_of(control).insert_paragraph("beside")
    assert out.text == "beside"
    assert blocks_of(control) == ["p"]


# ---------------------------------------------------------------------------
# what stays allowed (section 3.2): text into the one paragraph, formatting, delete, xml_mapping.delete()
# ---------------------------------------------------------------------------

ALLOWED = {
    "ContentControl.insert_text End": (lambda c: c.insert_text("X"), "NoteX"),
    "ContentControl.insert_text Start": (lambda c: c.insert_text("X", location="Start"), "XNote"),
    "ContentControl.insert_text Replace": (lambda c: c.insert_text("X", location="Replace"), "X"),
    "Body.insert_text Replace": (lambda c: c.body().insert_text("X", location="Replace"), "X"),
    "Paragraph.text": (lambda c: setattr(paragraph_of(c), "text", "X"), "X"),
    "Paragraph.insert_text End": (lambda c: paragraph_of(c).insert_text("X"), "NoteX"),
    "Paragraph.insert_text Replace": (lambda c: paragraph_of(c).insert_text("X", location="Replace"), "X"),
    "Paragraph.replace_text": (lambda c: paragraph_of(c).replace_text("Note", "X"), "X"),
    "Body.replace_text": (lambda c: c.body().replace_text("Note", "X"), "X"),
    "Range.insert_text": (lambda c: paragraph_of(c).get_range("Content").insert_text("X"), "X"),
    "Range.delete": (lambda c: paragraph_of(c).search("ot")[0].delete(), "Ne"),
    "Paragraph.insert_xml End, one": (lambda c: paragraph_of(c).insert_xml(ONE_PARAGRAPH, location="End"), "NoteX"),
    "Paragraph.insert_ooxml Start, one": (lambda c: paragraph_of(c).insert_ooxml(ONE_PARAGRAPH, location="Start"), "XNote"),
    "Paragraph.insert_markdown End, one": (lambda c: paragraph_of(c).insert_markdown("X", location="End"), "NoteX"),
    "Range.insert_ooxml, one": (lambda c: paragraph_of(c).get_range("Content").insert_ooxml(ONE_PARAGRAPH), "X"),
}


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("kind", KINDS, ids=list(KINDS))
@pytest.mark.parametrize("writer", ALLOWED, ids=list(ALLOWED))
def test_the_writes_section_3_2_allows_still_work_and_reach_the_node(writer, kind, tracked):
    package, control = value_control(KINDS[kind], tracked=tracked)
    write, text = ALLOWED[writer]
    write(control)
    assert blocks_of(control) == ["p"], "one paragraph, tracked or not"
    assert control.text == text
    assert control.xml_mapping.custom_xml_node.text == text, "CR-019 D7: every text write reaches the node"
    again = reloaded(package)
    assert len(again.body.content_controls[0].paragraphs) == 1


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("kind", SINGLE, ids=SINGLE)
def test_a_paragraph_break_in_a_single_line_one_is_dropped_by_every_text_writer(kind, tracked):
    """Section 3.6, check 56 cases 3, 11 and 15."""
    for write in (
        lambda c: c.insert_text("a\nb\rc", location="End"),
        lambda c: c.insert_text("a\nb\rc", location="Replace"),
        lambda c: paragraph_of(c).insert_text("a\nb\rc", location="End"),
        lambda c: paragraph_of(c).get_range("Content").insert_text("a\nb\rc"),
        lambda c: paragraph_of(c).get_range("End").insert_text("a\nb\rc", location="After"),
        lambda c: paragraph_of(c).search("Note")[0].insert_text("a\nb\rc", location="After"),
        lambda c: paragraph_of(c).replace_text("Note", "a\nb\rc"),
    ):
        _package, control = value_control(KINDS[kind], tracked=tracked)
        write(control)
        assert blocks_of(control) == ["p"]
        assert "abc" in control.text
        assert "\x0b" not in control.text


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_a_paragraph_break_in_a_multi_line_one_is_a_line_break(tracked):
    package, control = value_control(KINDS["w:text multi-line"], tracked=tracked)
    control.insert_text("a\nb", location="Replace")
    assert blocks_of(control) == ["p"]
    assert control.text == "a\x0bb"
    assert control.xml_mapping.custom_xml_node.text == "a\nb"
    assert "<w:br/>" in package.main_document_part.get_xml()


def test_formatting_stays_allowed():
    _package, control = value_control("")
    paragraph = paragraph_of(control)
    paragraph.font.bold = True
    paragraph.alignment = "Centered"
    assert blocks_of(control) == ["p"]


def test_deleting_the_control_keeps_or_takes_its_paragraph():
    package, control = value_control("")
    control.delete(keep_content=True)
    assert package.body.content_controls == []
    assert [p.text for p in package.body.paragraphs] == ["before", "Note", "after"]


def test_after_xml_mapping_delete_the_control_is_free():
    _package, control = value_control("")
    control.xml_mapping.delete()
    assert control.type == "RichText"
    control.insert_paragraph("two")
    control.body().insert_table(1, 1)
    assert blocks_of(control) == ["p", "p", "tbl"]


def test_into_a_control_emptied_a_paragraph_may_go_back():
    """One paragraph is what the rule asks for, and an empty control holds none yet."""
    _package, control = value_control("")
    paragraph_of(control).delete()
    assert blocks_of(control) == []
    control.body().insert_paragraph("back")
    assert blocks_of(control) == ["p"]
    with pytest.raises(BindingError):
        control.body().insert_paragraph("second")


# ---------------------------------------------------------------------------
# a Replace over a file made elsewhere (section 3.5): what is there isn't refused, and no write makes it worse
# ---------------------------------------------------------------------------


def test_a_paragraph_of_a_control_holding_two_replaced_as_a_block_is_refused():
    """What the write would leave is two paragraphs, core-ts's ``assertControlHolds`` with its ``replacing``."""
    package, control = holding_two()
    before = snapshot(package)
    with pytest.raises(BindingError):
        paragraph_of(control).insert_xml(ONE_PARAGRAPH, location="Replace")
    assert snapshot(package) == before


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_a_control_holding_two_refuses_a_third(tracked):
    package, control = holding_two()
    if tracked:
        package.change_tracking_mode = "TrackAll"
    before = snapshot(package)
    with pytest.raises(BindingError):
        paragraph_of(control).insert_paragraph("third")
    assert snapshot(package) == before


def test_a_control_around_the_paragraphs_of_a_control_holding_two_is_refused():
    """A range across its paragraphs, whole (``wrap_paragraphs``): the outer control would hold a control."""
    package, control = holding_two()
    span = control.body().get_range("Whole")
    assert span.spans_paragraphs
    before = snapshot(package)
    with pytest.raises(BindingError):
        span.insert_content_control()
    assert snapshot(package) == before


def test_a_control_holding_two_replaced_as_text_holds_one():
    """Across paragraphs (CR-013 Phase C) the deletion joins them; a break in the text makes none."""
    _package, control = holding_two()
    control.insert_text("x\ny", location="Replace")
    assert blocks_of(control) == ["p"]
    assert control.text == "xy"


def test_a_block_of_a_control_holding_two_can_be_removed():
    _package, control = holding_two()
    control.paragraphs[1].delete()
    assert [p.text for p in control.paragraphs] == ["first"]
    assert control.is_multi_line is False


def test_a_replace_of_the_only_paragraph_by_one_is_allowed_untracked_and_refused_tracked():
    """Tracked, the paragraph it replaces stays, deleted: that is two paragraphs."""
    package, control = value_control("")
    paragraph_of(control).insert_xml(ONE_PARAGRAPH, location="Replace")
    assert [p.text for p in control.paragraphs] == ["X"]
    package, control = value_control("", tracked=True)
    before = snapshot(package)
    with pytest.raises(BindingError):
        paragraph_of(control).insert_xml(ONE_PARAGRAPH, location="Replace")
    assert snapshot(package) == before


def test_a_body_replace_by_one_paragraph_is_allowed_untracked_and_refused_tracked():
    _package, control = value_control("")
    control.body().insert_ooxml(ONE_PARAGRAPH, location="Replace")
    assert [p.text for p in control.paragraphs] == ["X"]
    package, control = value_control("", tracked=True)
    before = snapshot(package)
    with pytest.raises(BindingError):
        control.body().insert_ooxml(ONE_PARAGRAPH, location="Replace")
    assert snapshot(package) == before


# ---------------------------------------------------------------------------
# set_mapping and set_mapping_by_node (section 3.3, Q4)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("kind", KINDS, ids=list(KINDS))
@pytest.mark.parametrize("by_node", [False, True], ids=["set_mapping", "set_mapping_by_node"])
def test_set_mapping_refuses_a_control_of_a_value_kind_holding_more(kind, by_node):
    package, control = holding_two(KINDS[kind], binding=False)
    part = package.custom_xml_parts.items[0]
    before = snapshot(package)
    with pytest.raises(BindingError) as refused:
        if by_node:
            control.xml_mapping.set_mapping_by_node(part.select_single_node(XPATH, MAPPINGS))
        else:
            control.xml_mapping.set_mapping(XPATH, MAPPINGS, part)
    assert refused.value.code == "binding.one_paragraph"
    assert "leave it unbound" in (refused.value.hint or "")
    assert snapshot(package) == before
    assert not control.xml_mapping.is_mapped


@pytest.mark.parametrize("kind", KINDS, ids=list(KINDS))
@pytest.mark.parametrize("by_node", [False, True], ids=["set_mapping", "set_mapping_by_node"])
def test_set_mapping_binds_a_control_of_one_paragraph(kind, by_node):
    package, control = holding_two(KINDS[kind], binding=False)
    part = package.custom_xml_parts.items[0]
    control.paragraphs[1].delete()
    if by_node:
        assert control.xml_mapping.set_mapping_by_node(part.select_single_node(XPATH, MAPPINGS))
    else:
        assert control.xml_mapping.set_mapping(XPATH, MAPPINGS, part)
    assert control.xml_mapping.is_mapped
    assert is_bound_block_control(control)


def test_set_mapping_refuses_a_table_too():
    package = create_package()
    package.body.insert_xml(control_xml("", content=A_TABLE.replace(f' xmlns:w="{W}"', "") + "<w:p/>"))
    control = package.body.content_controls[0]
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>v</value></data>')
    with pytest.raises(BindingError):
        control.xml_mapping.set_mapping(XPATH, MAPPINGS, part)


def test_a_path_that_selects_nothing_answers_false_before_any_refusal():
    package, control = holding_two(binding=False)
    assert control.xml_mapping.set_mapping("/ns0:data[1]/ns0:none[1]", MAPPINGS, package.custom_xml_parts.items[0]) is False


def test_a_mapping_refreshed_on_a_w15_binding_is_no_w_data_binding():
    """Word's rich-text binding may hold blocks: section 3.3 is ``w:dataBinding``'s."""
    package, control = holding_two(binding=False)
    from docx4j_py.w15 import el as w15_el

    part = package.custom_xml_parts.items[0]
    control.put_property(w15_el.dataBinding(xpath=XPATH, prefix_mappings=MAPPINGS, store_item_id=part.id))
    assert control.xml_mapping.set_mapping(XPATH, MAPPINGS, part)
    assert sdt_property(control.sdt_pr, "dataBinding", W15_NS) is not None
    assert [p.text for p in control.paragraphs] == ["first", "second"]


def test_a_date_control_holding_two_is_no_value_kind():
    """The date and list kinds are unmeasured (section 2.8): the rule leaves them."""
    package, control = holding_two('<w:date><w:dateFormat w:val="d MMMM yyyy"/></w:date>', binding=False)
    assert control.xml_mapping.set_mapping(XPATH, MAPPINGS, package.custom_xml_parts.items[0])
    assert not is_bound_block_control(control)
    assert control.type == "DatePicker"


# ---------------------------------------------------------------------------
# apply_bindings (section 3.4, Q5)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("kind", KINDS, ids=list(KINDS))
def test_apply_bindings_over_file_3s_shape_writes_one_paragraph_its_properties_kept(kind):
    package, control = holding_two(KINDS[kind])
    result = package.custom_xml_parts.apply_bindings()
    assert blocks_of(control) == ["p"]
    assert control.text == "bound value"
    assert paragraph_of(control).alignment == "Centered", "its w:pPr kept"
    (entry,) = result.applied
    assert [note.code for note in entry.notes] == ["binding.blocks_removed"]
    assert "1 other block" in entry.notes[0].message
    assert entry.to_dict()["notes"] == [{"code": "binding.blocks_removed", "message": entry.notes[0].message}]
    report = package.changes[-1]
    assert any("1 other block" in warning for warning in report.warnings)
    again = reloaded(package)
    assert [p.text for p in again.body.content_controls[0].paragraphs] == ["bound value"]


def test_apply_bindings_over_one_paragraph_says_nothing():
    package, _control = value_control("")
    (entry,) = package.custom_xml_parts.apply_bindings().applied
    assert entry.notes == ()
    assert "notes" not in entry.to_dict()


def test_apply_bindings_over_more_shows_its_placeholder_in_one_paragraph_for_an_empty_value():
    package, control = holding_two("", value="")
    package.custom_xml_parts.apply_bindings()
    assert blocks_of(control) == ["p"]
    assert control.is_showing_placeholder


def test_apply_bindings_takes_what_the_extra_blocks_held_with_them():
    """A comment whose reference went goes, as any deletion takes it (core-ts's ``removeBlocks``)."""
    package, control = holding_two("")
    control.paragraphs[1].get_range("Content").insert_comment("gone with it")
    assert len(package.body.get_comments()) == 1
    package.custom_xml_parts.apply_bindings()
    assert blocks_of(control) == ["p"]
    assert package.body.get_comments() == []
    assert package.main_document_part.comments_part is None, "settled as the pass ends, its last comment gone, the part with it"


def test_apply_bindings_leaves_markers_beside_the_paragraph():
    package = create_package()
    package.body.insert_xml(
        control_xml(
            "",
            content='<w:bookmarkStart w:id="1" w:name="b"/><w:p><w:r><w:t>first</w:t></w:r></w:p><w:bookmarkEnd w:id="1"/>'
            "<w:p><w:r><w:t>second</w:t></w:r></w:p>",
        )
    )
    control = package.body.content_controls[0]
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>v</value></data>')
    control.put_property(el.dataBinding(xpath=XPATH, prefix_mappings=MAPPINGS, store_item_id=part.id))
    package.custom_xml_parts.apply_bindings()
    assert blocks_of(control) == ["bookmarkStart", "p", "bookmarkEnd"]


def test_a_block_control_holding_a_table_stays_a_container():
    """Check 56's 56d: docx4j's binding pass and core-ts's leave it; Word couldn't open it anyway."""
    package = create_package()
    package.body.insert_xml(
        control_xml("", content=A_TABLE.replace(f' xmlns:w="{W}"', "") + "<w:p><w:r><w:t>n</w:t></w:r></w:p>")
    )
    control = package.body.content_controls[0]
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>v</value></data>')
    control.put_property(el.dataBinding(xpath=XPATH, prefix_mappings=MAPPINGS, store_item_id=part.id))
    assert is_container(control)
    assert not is_bound_block_control(control)
    assert control.type == "RichText"
    result = package.custom_xml_parts.apply_bindings()
    assert [entry.code for entry in result.skipped] == ["container"]
    assert blocks_of(control) == ["tbl", "p"]


def test_a_row_or_cell_control_keeps_the_first_paragraph_rule():
    """Section 3.4: CR-003 section 4's departure stays for a row or cell control."""
    package = create_package()
    package.body.insert_xml(
        f'<w:tbl xmlns:w="{W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:sdt><w:sdtPr><w:id w:val="3"/>'
        "</w:sdtPr><w:sdtContent><w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc>"
        "</w:sdtContent></w:sdt></w:tr></w:tbl>"
    )
    control = package.body.content_controls[0]
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>v</value></data>')
    control.put_property(el.dataBinding(xpath=XPATH, prefix_mappings=MAPPINGS, store_item_id=part.id))
    assert control.form == "cell"
    assert not is_bound_block_control(control)
    assert control.type == "RichText"
    assert value_control_of(control.element, package.body) is None


# ---------------------------------------------------------------------------
# the reading (section 3.5a, Q9)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("kind", "multi"),
    [("no type", False), ("w:text", False), ("w:text multi-line", True), ("w:richText", False)],
)
def test_a_value_control_reads_as_plain_text(kind, multi):
    package, control = value_control(KINDS[kind])
    assert control.type == "PlainText"
    assert control.is_multi_line is multi
    assert plain_text_lines_of(control) == ("multi" if multi else "single")
    assert control.to_dict()["type"] == "PlainText"
    assert "PlainText" in repr(control)
    assert package.body.get_content_controls(types=["PlainText"]) == [control]
    assert package.body.get_content_controls(types=["RichText"]) == []
    assert package.body.content_controls == [control]
    assert sdt_kind_of(control.sdt_pr) == ("PlainText" if "w:text" in kind else "RichText"), "the file as it is"


def test_the_reading_changes_nothing_in_the_file():
    package, control = value_control("")
    assert control.type == "PlainText"
    assert "w:text" not in re.sub(r"w:t>|w:t ", "", package.main_document_part.get_xml().split("sdtContent")[0])


def test_is_multi_line_is_what_apply_bindings_writes_a_newline_by():
    package, control = value_control("", value="x")
    control.xml_mapping.custom_xml_node.text = "one\ntwo"
    package.custom_xml_parts.apply_bindings()
    assert control.text == "onetwo", "Word drops the break"


def test_a_multi_line_value_control_keeps_the_break_apply_bindings_writes():
    package, control = value_control(KINDS["w:text multi-line"], value="x")
    control.xml_mapping.custom_xml_node.text = "one\ntwo"
    package.custom_xml_parts.apply_bindings()
    assert control.text == "one\x0btwo"


def test_a_binding_entry_reports_the_type_a_reader_does():
    package, _control = value_control("")
    (entry,) = package.custom_xml_parts.apply_bindings().applied
    assert entry.kind == "PlainText"


def test_describe_reads_the_same_with_the_part_read_and_not():
    package, _control = value_control("")
    data = package.save()
    assert [binding.kind for binding in package.custom_xml_parts.describe().bindings] == ["PlainText"]
    fresh = load(data)
    described = fresh.custom_xml_parts.describe()
    assert [binding.kind for binding in described.bindings] == ["PlainText"], "read as bytes"
    assert described.warnings


def test_describe_as_bytes_leaves_other_forms_as_they_are():
    package = create_package()
    package.body.insert_paragraph("run")
    package.body.paragraphs[0].search("run")[0].insert_content_control()
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>v</value></data>')
    package.body.content_controls[0].xml_mapping.set_mapping(XPATH, MAPPINGS, part)
    fresh = load(package.save())
    assert [binding.kind for binding in fresh.custom_xml_parts.describe().bindings] == ["RichText"]
    assert fresh.body.content_controls[0].type == "RichText", "section 3.5a is block-level (core-ts section 61.7)"


def test_to_api_script_writes_the_same_body_for_a_value_control():
    package, _control = value_control("")
    script = package.body.to_api_script()
    assert "body" in script
    assert "Note" in script


# ---------------------------------------------------------------------------
# Q9: a bound w:richText block control is bound as text, both ways
# ---------------------------------------------------------------------------


def test_apply_bindings_writes_a_bound_rich_text_block_control():
    package, control = value_control("<w:richText/>", value="x")
    control.xml_mapping.custom_xml_node.text = "from the data"
    result = package.custom_xml_parts.apply_bindings()
    assert [entry.kind for entry in result.applied] == ["PlainText"]
    assert control.text == "from the data"


def test_update_from_content_controls_writes_a_bound_rich_text_block_control():
    """CR-013 section 11.2: a control written ``NoteDxy`` showed ``NoteD`` in Word, its node's text, until this."""
    package, control = value_control("<w:richText/>")
    run = paragraph_of(control).runs[0]
    run.content[0].value = "edited"
    result = package.custom_xml_parts.update_from_content_controls()
    assert [entry.value for entry in result.applied] == ["edited"]
    assert control.xml_mapping.custom_xml_node.text == "edited"


def test_a_rich_text_control_written_reaches_its_node():
    _package, control = value_control("<w:richText/>")
    control.insert_text("Dxy")
    assert control.xml_mapping.custom_xml_node.text == "NoteDxy"


def rich_run_control():
    package = create_package()
    package.body.insert_paragraph("seed")
    package.body.paragraphs[0].insert_xml(
        f'<w:p xmlns:w="{W}"><w:sdt><w:sdtPr><w:id w:val="4"/><w:richText/></w:sdtPr><w:sdtContent>'
        "<w:r><w:t>Note</w:t></w:r></w:sdtContent></w:sdt></w:p>",
        location="End",
    )
    control = package.body.content_controls[0]
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>v</value></data>')
    control.put_property(el.dataBinding(xpath=XPATH, prefix_mappings=MAPPINGS, store_item_id=part.id))
    return package, control


def test_a_run_level_rich_text_control_stays_skipped():
    """Core-ts section 61.7 records Word binding one at run level too; not this CR's (open point)."""
    package, control = rich_run_control()
    assert control.form == "run"
    assert control.type == "RichText"
    assert [entry.code for entry in package.custom_xml_parts.apply_bindings().skipped] == ["rich_text"]


# ---------------------------------------------------------------------------
# what the rule leaves alone
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("kind", ["", "<w:richText/>"], ids=["no type", "w:richText"])
def test_an_unbound_control_is_unaffected(kind):
    package = create_package()
    package.body.insert_xml(control_xml(kind))
    control = package.body.content_controls[0]
    assert control.type == "RichText"
    assert control.is_multi_line is True
    control.insert_paragraph("two")
    control.body().insert_markdown("A\n\nB")
    assert [p.text for p in control.paragraphs] == ["Note", "two", "A", "B"]


def test_a_plain_text_control_unbound_is_unaffected():
    package = create_package()
    package.body.insert_xml(control_xml("<w:text/>"))
    control = package.body.content_controls[0]
    control.insert_paragraph("two")
    assert [p.text for p in control.paragraphs] == ["Note", "two"]


def test_a_w15_binding_is_left_alone():
    """``w15:dataBinding`` is Word's rich-text binding, not this (core-ts's rule too)."""
    package = create_package()
    package.body.insert_xml(control_xml(""))
    control = package.body.content_controls[0]
    from docx4j_py.w15 import el as w15_el

    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>v</value></data>')
    control.put_property(w15_el.dataBinding(xpath=XPATH, prefix_mappings=MAPPINGS, store_item_id=part.id))
    assert not is_bound_block_control(control)
    assert control.type == "RichText"
    control.insert_paragraph("two")
    assert len(control.paragraphs) == 2


def test_a_node_of_flat_opc_is_left_alone():
    """H1: Word's own rich-text binding, not measured."""
    package, control = value_control("<w:richText/>")
    control.xml_mapping.custom_xml_node.text = (
        '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"></pkg:package>'
    )
    assert not is_bound_block_control(control)
    assert control.type == "RichText"
    control.insert_paragraph("two")
    assert len(control.paragraphs) == 2
    assert [entry.code for entry in package.custom_xml_parts.apply_bindings().skipped] == ["rich_text"]


def test_a_container_is_left_alone():
    package = create_package()
    package.body.insert_xml(
        control_xml(
            "",
            content='<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>inner</w:t></w:r></w:p>'
            "</w:sdtContent></w:sdt>",
        )
    )
    control = package.body.content_controls[0]
    part = package.custom_xml_parts.add(f'<data xmlns="{DATA}"><value>v</value></data>')
    control.put_property(el.dataBinding(xpath=XPATH, prefix_mappings=MAPPINGS, store_item_id=part.id))
    assert is_container(control)
    assert not is_bound_block_control(control)
    assert control.type == "RichText"


def test_a_control_of_another_kind_is_left_alone():
    _package, control = holding_two('<w:comboBox><w:listItem w:displayText="a" w:value="a"/></w:comboBox>')
    assert control.type == "ComboBox"
    assert not is_bound_block_control(control)
    assert holds_one_paragraph(control) is False
    assert blocks_of(control) == ["p", "p"]


def test_a_control_in_invoice2013_is_read_as_word_reads_it():
    """Word's own file: every bound block control of a value kind reads as plain text, a w15 one as it is."""
    package = sample("invoice2013.docx")
    for control in package.body.content_controls:
        if is_bound_block_control(control):
            assert control.type == "PlainText"
        elif control.form != "block":
            assert control.type == sdt_kind_of(control.sdt_pr)


def test_a_file_holding_such_a_control_is_read_and_saved_unchanged():
    """Section 3.5: CLAUDE.md's fidelity rule and core-ts's D4; nothing here repairs on load or save."""
    package, _control = holding_two("")
    data = package.save()
    fresh = load(data)
    assert fresh.main_document_part.get_xml() == load(data).main_document_part.get_xml()
    control = fresh.body.content_controls[0]
    assert [p.text for p in control.paragraphs] == ["first", "second"]
    fresh.custom_xml_parts.update_from_content_controls()
    assert control.xml_mapping.custom_xml_node.text == "first\nsecond", "F1: its paragraphs a \\n apart"
    assert [p.text for p in control.paragraphs] == ["first", "second"]
    assert [p.text for p in reloaded(fresh).body.content_controls[0].paragraphs] == ["first", "second"]
