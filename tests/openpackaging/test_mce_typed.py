"""A kept ``mc:AlternateContent`` branch's children typed: CR-012 Phase B, docx4j's 61 roots.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_mce_typed.py -q

JAXB binds a lax wildcard's child by its ``@XmlRootElement`` class, so in
docx4j a ``w:drawing`` in an ``mc:Choice`` is a ``Drawing``. This model typed a
wildcard's child by its global elements only, and 61 WordprocessingML names
that are roots in docx4j are not global here. Phase B gives them to the fork
parser's branch hook (``docx4j_py.runtime.branch_child_class``), from a table
derived from docx4j's own classes (``codegen/names/root_elements.json``),
generated into ``docx4j_py.el_index.ROOT_ELEMENTS``.

What is held (CR-012 section 5, Phase B):

* the table: derived, generated and registered alike, 61 names, ``w:sdtContent``
  left out (Q3), each the class the ``el`` table builds;
* each of the 61 resolves through the parser to its class;
* the hook's three steps (section 4.2 item 1): (a) in a branch, the container's
  binding, which beats (b) the root class, in any other lax wildcard and in a
  branch whose container does not take the name, and (c) None, the lookup by
  name;
* a ``w:drawing`` in an ``mc:Choice`` and a ``w:pict`` in an ``mc:Fallback``,
  under a ``w:r``, typed and saved whole, with no ``xsi:type``;
* the serialiser's element names a function, class to name, naming every
  class step (a) can give; ``to_xml`` unchanged for each, with a name or
  without;
* the census on ``samples/``, pinned;
* the walks name a branch's typed child as the same element elsewhere;
* the readers this phase moved (section 11): ``Body.tables`` through a
  block-level branch, the python-docx facade's inline shapes, a paragraph in a
  branch one line of ``text_of``;
* a ``word``-mode record's parse types the roots too.
"""

from __future__ import annotations

import dataclasses
import importlib
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon
import mce_branches

from docx4j_py import (
    load,
    runtime,
)
from docx4j_py.el_index import (
    EL_MODULES,
    ROOT_ELEMENT_BINDINGS,
    ROOT_ELEMENTS,
)
from docx4j_py.fragments import to_xml, wml
from docx4j_py.mce import (
    AlternateContent,
    AlternateContentChoice,
    AlternateContentFallback,
)
from docx4j_py.model.content.errors import AddressError
from docx4j_py.traversal import find, text_of, walk
from docx4j_py.wml import (
    RT,
    Body1,
    CtRTab,
    CTSdtCell,
    CTSdtRow,
    CTSimpleField,
    CtSimpleFieldHyperlink,
    Drawing,
    P,
    PHyperlink,
    Pict,
    R,
    RTab,
    SdtBlock,
    SdtRun,
    T,
    Tbl,
    Tc,
    Tr,
)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
SAMPLES = ROOT / "samples"
FIXTURES = ROOT / "tests" / "fixtures"
TABLE = json.loads((ROOT / "codegen" / "names" / "root_elements.json").read_text(encoding="utf-8"))


def w(local: str) -> str:
    return f"{{{_W}}}{local}"


def _class(path: str) -> type:
    module, _, name = path.rpartition(".")
    return getattr(importlib.import_module(module), name)


def choice(ac: AlternateContent) -> list:
    return list(ac.choice[0].any_element)


# ---------------------------------------------------------------------------
# the table
# ---------------------------------------------------------------------------


def test_the_table_is_derived_generated_and_registered_alike():
    """root_elements.json (derive_names.py) is el_index.ROOT_ELEMENTS (generate_el.py) is the runtime's."""
    assert TABLE["namespace"] == _W
    assert len(TABLE["roots"]) == 61
    assert sorted(ROOT_ELEMENTS) == sorted(w(local) for local in TABLE["roots"])
    for local, entry in TABLE["roots"].items():
        assert ROOT_ELEMENTS[w(local)] == f"docx4j_py.wml.{entry['class']}"
        assert runtime.root_element_class(w(local)) is _class(ROOT_ELEMENTS[w(local)])
    assert runtime.wildcard_names() == runtime.BRANCH_CHILD_NAMES | frozenset(ROOT_ELEMENTS)


def test_the_table_leaves_sdt_content_out_and_keeps_sdt_pr():
    """Q3: w:sdtContent has four classes; w:sdtPr is the 62nd root, annotated spaced in the schema copy."""
    assert w("sdtContent") not in ROOT_ELEMENTS
    assert "sdtContent" in TABLE["left_out"]
    assert ROOT_ELEMENTS[w("sdtPr")] == "docx4j_py.wml.SdtPr"
    assert TABLE["cross_check"] == {
        "schema": "schemas/wml/wml.xsd",
        "annotated_and_not_global": 62,
        "agree": True,
    }


def test_the_two_class_choices_section_2_states():
    """w:pPr and friends take docx4j's root classes, not the ...Base types; w:spacing PPrBase.Spacing's."""
    assert ROOT_ELEMENTS[w("pPr")] == "docx4j_py.wml.PPr"
    assert ROOT_ELEMENTS[w("sectPr")] == "docx4j_py.wml.SectPr"
    assert ROOT_ELEMENTS[w("tblPr")] == "docx4j_py.wml.TblPr"
    assert ROOT_ELEMENTS[w("tblGrid")] == "docx4j_py.wml.TblGrid"
    assert ROOT_ELEMENTS[w("spacing")] == "docx4j_py.wml.CtPprBaseSpacing"
    assert TABLE["roots"]["spacing"]["docx4j"] == "PPrBase.Spacing"


@pytest.mark.parametrize("qname", sorted(ROOT_ELEMENTS), ids=lambda q: q.rpartition("}")[2])
def test_each_root_is_the_class_the_el_table_builds(qname):
    from docx4j_py.wml.el import QNAME_TO_CLASS

    assert QNAME_TO_CLASS[qname] is _class(ROOT_ELEMENTS[qname])
    assert runtime.context().find_type(qname) is None  # not global here: that is why it is a root


@pytest.mark.parametrize("qname", sorted(ROOT_ELEMENTS), ids=lambda q: q.rpartition("}")[2])
def test_each_root_resolves_through_the_parser_to_its_class(qname):
    """No container above the element: (b), the root class; saved back under its name, no xsi:type."""
    local = qname.rpartition("}")[2]
    ac = wml(f'<mc:AlternateContent><mc:Choice Requires="w14"><w:{local}/></mc:Choice></mc:AlternateContent>')
    (child,) = choice(ac)
    assert type(child) is _class(ROOT_ELEMENTS[qname])
    xml = to_xml(ac)
    assert f"<w:{local}/>" in xml
    assert "xsi:type" not in xml


# ---------------------------------------------------------------------------
# the hook's three steps
# ---------------------------------------------------------------------------

BRANCH = (AlternateContentChoice, AlternateContent)


def test_a_containers_binding_beats_the_root_class_in_a_branch():
    """(a): a simple field binds w:hyperlink to its own class, a math run w:t and w:tab to its own."""
    assert runtime.branch_child_class(w("hyperlink"), (*BRANCH, CTSimpleField)) is CtSimpleFieldHyperlink
    assert runtime.branch_child_class(w("hyperlink"), (*BRANCH, P)) is PHyperlink
    field = wml(
        '<w:fldSimple w:instr=" X "><mc:AlternateContent><mc:Choice Requires="w14">'
        '<w:hyperlink w:anchor="a"><w:r><w:t>x</w:t></w:r></w:hyperlink>'
        "</mc:Choice></mc:AlternateContent></w:fldSimple>"
    )
    (link,) = choice(field.content[0])
    assert type(link) is CtSimpleFieldHyperlink
    assert '<w:hyperlink w:anchor="a"><w:r><w:t>x</w:t></w:r></w:hyperlink>' in to_xml(field)

    math_run = wml('<m:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>x</w:t><w:tab/></mc:Choice></mc:AlternateContent></m:r>')
    assert [type(c) for c in choice(math_run.content[0])] == [T, CtRTab]
    run = wml('<w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>x</w:t><w:tab/></mc:Choice></mc:AlternateContent></w:r>')
    assert [type(c) for c in choice(run.content[0])] == [RT, RTab]


def test_the_root_class_where_the_container_does_not_take_the_name():
    """(b) in a branch: a body takes no w:hyperlink, so its branch's is the root class."""
    assert runtime.container_binding(w("hyperlink"), (*BRANCH, Body1)) is None
    assert runtime.branch_child_class(w("hyperlink"), (*BRANCH, Body1)) is PHyperlink
    assert runtime.branch_child_class(w("tbl"), (*BRANCH, Body1)) is Tbl
    assert runtime.branch_child_class(w("tbl"), (*BRANCH, P)) is Tbl  # a paragraph takes no table


def test_the_root_class_in_a_lax_wildcard_that_is_not_a_branch():
    """(b) in any lax wildcard, as JAXB's lax binding; (c) the lookup by name for a name with no root."""
    from docx4j_py.dml.main import GraphicData

    assert runtime.branch_child_class(w("tc"), (GraphicData,)) is Tc
    assert runtime.branch_child_class(w("sdt"), (GraphicData,)) is None
    data = wml(
        '<a:graphicData uri="urn:x"><w:tbl/><w:tc/><w:sdt/><w:tr/><w:ins/>'
        "<w:instrText>i</w:instrText></a:graphicData>"
    )
    kinds = [type(c).__name__ for c in data.any_element]
    # tbl and tc by (b); sdt and tr by the lookup, (c); ins global; instrText neither: a wildcard
    assert kinds == ["Tbl", "Tc", "SdtBlock", "Tr", "RunIns", "AnyElement"]
    assert to_xml(data).endswith("<w:tbl/><w:tc/><w:sdt/><w:tr/><w:ins/><w:instrText>i</w:instrText></a:graphicData>")


def test_no_other_name_is_typed_in_a_branch():
    """A branch's w:instrText, w:moveFrom and w:customXml stay wildcard trees, as docx4j keeps them DOM elements."""
    from docx4j_py.child import is_any_element

    paragraph = wml(
        '<w:p><mc:AlternateContent><mc:Choice Requires="w14">'
        '<w:moveFrom w:id="1" w:author="a"><w:r><w:t>m</w:t></w:r></w:moveFrom>'
        '<w:customXml w:element="x"><w:r><w:t>c</w:t></w:r></w:customXml>'
        "</mc:Choice></mc:AlternateContent></w:p>"
    )
    assert all(is_any_element(c) for c in choice(paragraph.content[0]))
    assert w("moveFrom") not in runtime.wildcard_names()
    assert w("instrText") not in runtime.wildcard_names()


# ---------------------------------------------------------------------------
# a w:drawing in a Choice and a w:pict in a Fallback, under a w:r
# ---------------------------------------------------------------------------


def test_a_drawing_in_a_choice_and_a_pict_in_a_fallback_parse_typed_and_save_whole():
    pkg = load(SAMPLES / "DrawingML_GraphicData_wps.docx")
    part = pkg.main_document_part
    source = part.xml
    contents = part.contents
    (ac,) = find(contents, AlternateContent, mce="all")
    assert isinstance(ac.parent, R)
    assert [type(c) for c in choice(ac)] == [Drawing]
    assert [type(c) for c in ac.fallback.any_element] == [Pict]
    saved = part._marshal(contents)
    report = canon.compare(source, saved)
    assert report.identical, report.differences
    assert b"xsi:type" not in saved
    assert part.skipped == []


def test_the_choice_is_typed_all_the_way_down_and_the_fallback_to_its_pict():
    """Section 3's depth: 63 wildcard nodes become typed in the Choice; VML stays below the Pict (Phase C)."""
    from docx4j_py.wml import CTTxbxContent
    from docx4j_py.word.wps import Wsp

    contents = load(SAMPLES / "DrawingML_GraphicData_wps.docx").main_document_part.contents
    (ac,) = find(contents, AlternateContent, mce="all")
    counts = mce_branches.census(contents)
    assert counts["deep"][("choice", "any")] == 0
    assert counts["deep"][("choice", "typed")] > 0
    (drawing,) = choice(ac)
    assert find(drawing, Wsp) and find(drawing, CTTxbxContent)
    assert {name for (side, name) in counts["wildcard"] if side == "fallback"} >= {"rect", "textbox", "txbxContent"}


# ---------------------------------------------------------------------------
# saving: the element names
# ---------------------------------------------------------------------------


def test_the_element_names_are_a_function_naming_every_class_the_hook_can_give():
    names = runtime._branch_element_names()
    for qname, path in ROOT_ELEMENTS.items():
        assert names[_class(path)] == qname
    for qname, paths in ROOT_ELEMENT_BINDINGS.items():
        for path in paths:
            assert names[_class(path)] == qname
    for cls in (SdtBlock, SdtRun, CTSdtRow, CTSdtCell):
        assert names[cls] == w("sdt")
    assert names[Tr] == w("tr") and names[Tc] == w("tc")
    # 4 w:sdt classes, Tr, the 61 roots (Tc among them) and the 47 other bindings
    assert len(names) == 4 + 1 + 61 + 47
    assert set(names.values()) <= runtime.wildcard_names()


def test_a_class_claimed_by_two_names_is_refused():
    with pytest.raises(ValueError, match="claimed by"):
        runtime.element_names_table([(Tc, w("tc")), (Tc, w("tbl"))])
    assert runtime.element_names_table([(Tc, w("tc")), (Tc, w("tc"))]) == {Tc: w("tc")}


def _containers() -> list[type]:
    """Every class that can hold an mc:AlternateContent, by a field or a lax wildcard: step (a)'s containers."""
    out = []
    for module_name in sorted({m.rsplit(".", 1)[0] for m in EL_MODULES.values()}):
        module = importlib.import_module(module_name)
        for value in vars(module).values():
            if not (isinstance(value, type) and dataclasses.is_dataclass(value) and value.__module__ == module_name):
                continue
            if runtime._is_mc_wrapper(value):
                continue
            try:
                meta = runtime.context().build(value)
            except Exception:  # noqa: BLE001, S112 - inkml's MappingType has no metadata; generate_el.py passes it by too
                continue
            for var in meta.get_element_vars():
                types = [*var.types, *(t for alt in (var.elements or {}).values() for t in alt.types)]
                if (var.is_wildcard and var.process_contents != "skip") or AlternateContent in types:
                    out.append(value)
                    break
    return out


def test_every_class_step_a_can_give_is_named_under_that_name():
    """The generator's list, recomputed: for each container and root, the class it binds the name to is named."""
    names = runtime._branch_element_names()
    containers = _containers()
    assert Body1 in containers and R in containers and CTSimpleField in containers
    for qname in ROOT_ELEMENTS:
        for container in containers:
            found = runtime.container_binding(qname, (*BRANCH, container))
            if found is not None:
                assert names.get(found) == qname, (container.__name__, qname, found.__name__)


def _phase_a2_names() -> dict[type, str]:
    pairs = [(cls, w("sdt")) for cls in (SdtBlock, SdtRun, CTSdtRow, CTSdtCell)]
    return runtime.element_names_table([*pairs, (Tr, w("tr")), (Tc, w("tc"))])


@pytest.mark.parametrize("cls", sorted(set(runtime._branch_element_names()), key=lambda c: c.__name__), ids=lambda c: c.__name__)
def test_to_xml_is_unchanged_for_each_named_class(cls, monkeypatch):
    """The root is written under the table's name now, then renamed by to_xml as before: the same fragment."""
    obj = cls()
    qname = runtime._branch_element_names()[cls]
    now = to_xml(obj, name=qname)
    named = to_xml(obj, name=w("custom"))
    monkeypatch.setattr(runtime, "_branch_names", _phase_a2_names())
    assert to_xml(obj, name=qname) == now
    assert to_xml(obj, name=w("custom")) == named


def test_nothing_written_through_a_field_changes():
    """A part with no mc:AlternateContent saves byte for byte as with Phase A2's parser and names."""
    pkg = load(SAMPLES / "Symbols.docx")
    part = pkg.main_document_part
    source = part.xml
    assert b"AlternateContent" not in source
    typed = part._marshal(part.contents)
    with mce_branches.untyped_roots():
        again = load(SAMPLES / "Symbols.docx").main_document_part
        untyped = again._marshal(again.contents)
    assert typed == untyped


# ---------------------------------------------------------------------------
# the census on samples/, pinned
# ---------------------------------------------------------------------------


def test_the_census_on_samples():
    """Section 5: 3 Drawing in Choices, 3 Pict in Fallbacks, no wildcard w:drawing or w:pict left in a branch."""
    children: dict = {}
    deep: dict = {}
    for path in sorted(SAMPLES.glob("*.docx")):
        pkg = load(path)
        for _name, part in mce_branches.wml_parts(pkg):
            counts = mce_branches.census(part.contents)
            for key, value in counts["children"].items():
                children[key] = children.get(key, 0) + value
            for key, value in counts["deep"].items():
                deep[key] = deep.get(key, 0) + value
    assert children == {("choice", "typed:Drawing"): 3, ("fallback", "typed:Pict"): 3}
    assert deep == {("choice", "typed"): 185, ("fallback", "typed"): 3, ("fallback", "any"): 28}


# ---------------------------------------------------------------------------
# the walks
# ---------------------------------------------------------------------------


def test_walk_names_a_branchs_typed_child_as_elsewhere():
    contents = load(SAMPLES / "DrawingML_GraphicData_wps.docx").main_document_part.contents
    seen: list[tuple[str | None, str]] = []
    walk(contents, lambda node, parent, name: seen.append((name, type(node).__name__)), mce="all")
    assert (w("drawing"), "Drawing") in seen
    assert (w("pict"), "Pict") in seen
    assert not any(name == "any_element" for name, _ in seen)


@pytest.mark.parametrize(
    ("mode", "expected"), [("resolve", "a-choice"), ("all", "a-choice-fallback"), ("none", "a")]
)
def test_text_of_reads_a_paragraph_level_branchs_runs(mode, expected):
    """A branch's typed w:r was yielded under the field's name, any_element, and read by no walk of text."""
    paragraph = wml(
        "<w:p><w:r><w:t>a</w:t></w:r><mc:AlternateContent>"
        '<mc:Choice Requires="w14"><w:r><w:t>-choice</w:t></w:r></mc:Choice>'
        "<mc:Fallback><w:r><w:t>-fallback</w:t></w:r></mc:Fallback>"
        "</mc:AlternateContent></w:p>"
    )
    assert text_of(paragraph, mce=mode) == expected


def test_a_paragraph_in_a_block_level_branch_is_one_line():
    """mc-txbx-tc-block's cell: its branch paragraph's two runs read as one line, not two."""
    pkg = load(FIXTURES / "mc-fun" / "mc-txbx-tc-block.docx")
    (table,) = pkg.body.tables
    assert table.values == [["I prefer", "CELL_CHOICE_PARA "]]
    assert table.text == "I prefer\nCELL_CHOICE_PARA "


# ---------------------------------------------------------------------------
# the readers this phase moved
# ---------------------------------------------------------------------------

_BLOCK_TABLE = (
    "<w:body><w:p><w:r><w:t>before</w:t></w:r></w:p>"
    '<mc:AlternateContent><mc:Choice Requires="w14"><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>'
    "<w:tr><w:tc><w:p><w:r><w:t>CHOICE_CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl></mc:Choice>"
    "<mc:Fallback><w:p><w:r><w:t>FALLBACK</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>"
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
    "<w:sectPr/></w:body>"
)


def _block_table_package():
    from docx4j_py import create_package
    from docx4j_py.child import link_parents

    pkg = create_package()
    pkg.main_document_part.contents.body = wml(_BLOCK_TABLE)
    link_parents(pkg.main_document_part.contents)
    return pkg


def test_body_tables_reads_a_table_in_a_block_level_branch():
    """CR-012 section 9.11: a typed Tbl in Word's branch is one of the body's own tables, read-only."""
    body = _block_table_package().body
    tables = body.tables
    assert [(t.address, t.values) for t in tables] == [("body/1/0", [["CHOICE_CELL"]]), ("body/2", [[""]])]
    assert body.element_at("body/1/0").values == [["CHOICE_CELL"]]
    with pytest.raises(AddressError) as raised:
        tables[0].values = [["edited"]]
    assert raised.value.code == "address.read_only"
    with pytest.raises(AddressError):
        tables[0].add_rows(1)
    with pytest.raises(AddressError):
        tables[0].delete()
    assert tables[0].values == [["CHOICE_CELL"]]


def test_the_content_api_reads_a_block_level_branchs_table_as_its_paragraphs():
    """Body.paragraphs, outline() and get_text() reach the branch table's paragraphs; before, its w:tbl was a wildcard."""
    body = _block_table_package().body
    assert [(p.address, p.text) for p in body.paragraphs] == [
        ("body/0", "before"),
        ("body/1/0/0/0/0", "CHOICE_CELL"),
        ("body/2/0/0/0", ""),
    ]
    entry = body.outline().to_dict()["entries"][1]
    assert entry["kind"] == "AlternateContent"
    assert [(c["address"], c["kind"], c["text"]) for c in entry["children"]] == [("body/1/0", "table", "CHOICE_CELL")]
    assert "CHOICE_CELL" in body.get_text()


def test_the_python_docx_facade_keeps_a_drawing_straight_in_a_branch_out():
    """python-docx's inline shapes are //w:p/w:r/w:drawing/wp:inline: a branch's drawing is no run's child."""
    from docx4j_py.dml.wordprocessing_drawing import Inline
    from docx4j_py.docx import Document

    contents = load(FIXTURES / "waterfall.docx").main_document_part.contents
    (ac,) = find(contents, AlternateContent, mce="all")
    (drawing,) = choice(ac)
    assert isinstance(drawing, Drawing) and isinstance(drawing.anchor_or_inline[0], Inline)
    assert len(Document(str(FIXTURES / "waterfall.docx")).inline_shapes) == 0


def test_the_facade_still_lists_a_drawing_in_a_run():
    from docx4j_py.docx import Document

    shapes = Document(str(FIXTURES / "text-extraction" / "phase0" / "extraction-probe-inspect.docx")).inline_shapes
    assert len(shapes) == 2


# ---------------------------------------------------------------------------
# word mode: the record's own parse
# ---------------------------------------------------------------------------


def test_a_word_mode_records_parse_types_the_roots():
    """_record_hook: (a) with the level's container, then (b); a block-level record's w:tbl is a Tbl."""
    from docx4j_py.openpackaging.mc_block_level import (
        BLOCK,
        CELL,
        _record_hook,
        splice_word,
    )

    block = _record_hook(BLOCK)
    assert block(w("tbl"), BRANCH) is Tbl
    assert block(w("hyperlink"), BRANCH) is PHyperlink  # a body takes none: (b)
    assert block(w("sdt"), BRANCH) is SdtBlock
    assert _record_hook(CELL)(w("tc"), BRANCH) is Tc

    data = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        + _BLOCK_TABLE
        + "</w:document>"
    ).encode()
    _bytes, (record,) = splice_word(data)
    assert [type(c) for c in record.element.choice[0].any_element] == [Tbl]
    assert [type(c) for c in record.element.fallback.any_element] == [P]
    assert isinstance(record.element.fallback, AlternateContentFallback)


# ---------------------------------------------------------------------------
# pinned for the coordinator's decision (CR-012 section 11.4)
# ---------------------------------------------------------------------------

_CELLS = (
    '<w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>'
    "<w:tc><w:p><w:r><w:t>a</w:t></w:r><mc:AlternateContent>"
    '<mc:Choice Requires="w14"><w:r><w:t>-PARA_CHOICE</w:t></w:r></mc:Choice>'
    "<mc:Fallback><w:r><w:t>-PARA_FALLBACK</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p></w:tc>"
    "<w:tc><w:p><w:r><w:t>b</w:t><mc:AlternateContent>"
    '<mc:Choice Requires="w14"><w:t>-RUN_CHOICE</w:t></mc:Choice>'
    "<mc:Fallback><w:t>-RUN_FALLBACK</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p></w:tc>"
    "</w:tr></w:tbl><w:sectPr/></w:body>"
)


def test_pinned_for_the_coordinators_decision_text_of_readers_read_a_paragraph_level_branchs_runs():
    """The text_of readers now read Word's branch at paragraph level, as at run level; Paragraph.text reads neither.

    Before Phase B, ``text_of`` (``Block.text``, ``Table.text``, ``TableCell.text``,
    a run control's text, ``TrackedChange.text``) read a run-level branch's
    wildcard ``w:t`` but not a paragraph-level branch's typed ``w:r``, which the
    walks named ``any_element``: this cell read ``a``. Now it reads ``a-PARA_CHOICE``,
    as ``text_of``'s own contract says (the resolved branch contributes) and as
    Word shows. ``Paragraph.text`` and ``get_text``, Office JS's text by CR-011
    Q4, read neither branch, before and after. Whether the ``text_of`` readers
    that claim Office JS's text should descend at all is CR-011 Q4's question,
    left to the coordinator (CR-012 section 11.4).
    """
    from docx4j_py import create_package
    from docx4j_py.child import link_parents

    pkg = create_package()
    pkg.main_document_part.contents.body = wml(_CELLS)
    link_parents(pkg.main_document_part.contents)
    body = pkg.body
    (table,) = body.tables
    assert table.values == [["a-PARA_CHOICE", "b-RUN_CHOICE"]]
    assert table.text == "a-PARA_CHOICE\nb-RUN_CHOICE"
    assert [cell.text for cell in table.rows[0].cells] == ["a-PARA_CHOICE", "b-RUN_CHOICE"]
    assert [p.text for p in body.paragraphs] == ["a", "b"]
    assert body.get_text() == "a\nb"
