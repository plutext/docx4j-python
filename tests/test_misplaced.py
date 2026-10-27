"""Content refused where it does not belong, as code adds it. CR-020 Phase A (sections 4.1 to 4.6, 4.8 to 4.10, 5).

Every typed list refuses an item its field's element choices do not allow
(D2): a ``Tbl`` in ``R.content``, a ``str`` in ``P.content``, a ``Tc`` in
``Tbl.content``, with :class:`MisplacedContentError` (``target.misplaced``), a
message naming the holder's field and the item, and a hint saying where the
item goes (D6). The parse is not checked (D4), nor is anything under
:func:`docx4j_py.child.unchecked` (D9).

    .venv-fork/bin/python -m pytest tests/test_misplaced.py -q
"""

from __future__ import annotations

import ast
import dataclasses
import threading
from importlib import import_module
from pathlib import Path

import pytest
from lxml import etree

import docx4j_py.wml as W
from docx4j_py import placement
from docx4j_py.child import (
    ChildList,
    _list_field_names,
    deep_copy,
    link_parents,
    unchecked,
)
from docx4j_py.el_index import EL_MODULES
from docx4j_py.fragments import to_xml, wml
from docx4j_py.model.content import (
    ContentError,
    Docx4JError,
    InvalidTargetError,
    MisplacedContentError,
)
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.runtime import context, serializer
from docx4j_py.wml import el, p, r, t, tc, tr
from docx4j_py.wml.sdt import sdt

ROOT = Path(__file__).resolve().parent.parent
#: The 16 documents of ``samples/``, as ``openpackaging/test_round_trip.py`` takes them.
SAMPLES = sorted(set((ROOT / "samples").glob("*.doc*")) | set((ROOT / "samples").glob("*.dot*"))
                 | set((ROOT / "samples").glob("*.pptx")) | set((ROOT / "samples").glob("*.xlsx")))
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def refused(holder, field, item) -> MisplacedContentError:
    """Append `item` to `holder`'s `field`, which must refuse it; the list must not change."""
    items = getattr(holder, field)
    before = list(items)
    with pytest.raises(MisplacedContentError) as raised:
        items.append(item)
    assert list(items) == before
    return raised.value


# ---------------------------------------------------------------------------
# the error (D5)
# ---------------------------------------------------------------------------


def test_the_error_is_a_type_error_an_invalid_target_and_a_content_error():
    error = refused(r("x"), "content", W.Tbl())
    assert isinstance(error, TypeError)
    assert isinstance(error, InvalidTargetError)
    assert isinstance(error, ContentError)
    assert isinstance(error, Docx4JException)
    assert Docx4JError is Docx4JException
    assert error.code == "target.misplaced"
    assert error.to_dict() == {"code": "target.misplaced", "message": error.message, "hint": error.hint}
    assert error.message == "`R.content` can't hold a `Tbl` (`w:tbl`)"
    assert str(error) == f"{error.message} ({error.hint})"


def test_the_error_carries_the_holder_the_field_the_item_and_its_name():
    run = r("x")
    table = W.Tbl()
    error = refused(run, "content", table)
    assert error.holder is run
    assert error.field == "content"
    assert error.item is table
    assert error.element == f"{{{W_NS}}}tbl"


def test_the_error_is_exported_where_the_others_are():
    from docx4j_py.model.content import errors

    assert errors.MisplacedContentError is MisplacedContentError
    assert "MisplacedContentError" in errors.__all__
    assert MisplacedContentError.default_code == "target.misplaced"


# ---------------------------------------------------------------------------
# docx4j-25's six mistakes (CR-020 sections 1 and 2.7), each with its hint
# ---------------------------------------------------------------------------

SIX = [
    # 1. a run or bare text in the body or a cell
    ("1 run in body", W.Body, "content", lambda: r("x"), "a run goes in a paragraph: `p(r(...))`"),
    ("1 run in cell", W.Tc, "content", lambda: r("x"), "a run goes in a paragraph: `p(r(...))`"),
    ("1 text in cell", W.Tc, "content", lambda: t("x"), 'text goes in a run, in a paragraph: `p(r("..."))`'),
    # 2. a drawing straight in a paragraph
    ("2 drawing in p", W.P, "content", lambda: el.drawing(), "a drawing goes in a run: `el.r(content=[drawing])`"),
    # 3. a table or text in a paragraph (and a table in a run)
    ("3 table in p", W.P, "content", W.Tbl, placement._TABLE_SIBLING),
    ("3 table in run", W.R, "content", W.Tbl, placement._TABLE_SIBLING),
    ("3 text in p", W.P, "content", lambda: t("x"), 'text goes in a run: `r("...")`'),
    # 4. a plain string in a content list
    ("4 string in p", W.P, "content", lambda: "hello", 'a string is not an element: wrap it, `r("...")`'),
    ("4 string in run", W.R, "content", lambda: "hello", 'a string is not an element: wrap it, `t("...")`'),
    # 5. a property object as content
    ("5 pPr in p", W.P, "content", el.pPr, "`pPr` is a property, not content: set `p.p_pr`"),
    (
        "5 sectPr in body",
        W.Body,
        "content",
        el.sectPr,
        "`sectPr` is a property, not content: set `body.sect_pr`, or `p_pr.sect_pr` on the paragraph that ends the section",
    ),
    # 6. a level skipped
    ("6 p in tr", W.Tr, "content", lambda: p("x"), "a paragraph goes in a cell: `el.tc(content=[p(...)])`"),
    ("6 tc in tbl", W.Tbl, "content", lambda: tc("x"), "a table holds rows: `el.tr(content=[el.tc(...)])`"),
]


@pytest.mark.parametrize("case,holder,field,item,hint", SIX, ids=[case for case, *_ in SIX])
def test_docx4j_25s_six_mistakes_are_refused_with_the_fix(case, holder, field, item, hint):
    error = refused(holder(), field, item())
    assert error.code == "target.misplaced"
    assert error.hint == hint
    assert error.message.startswith(f"`{holder.__name__}.{field}` can't hold ")


def test_the_message_names_both_the_class_and_the_element():
    assert refused(W.P(), "content", t("x")).message == "`P.content` can't hold an `RT` (`w:t`)"
    assert refused(W.Body(), "content", r("x")).message == "`Body.content` can't hold an `R` (`w:r`)"
    assert refused(W.P(), "content", "hello").message == "`P.content` can't hold a `str`"


# ---------------------------------------------------------------------------
# Java's cases: ArrayListDocx4j.checkParent, and docx4j CR-064 section 2's table
# ---------------------------------------------------------------------------

JAVA = [
    # P and P.Hyperlink (since 49637c6e4): a Tbl, a P, bare Text
    ("P tbl", W.P, W.Tbl, placement._TABLE_SIBLING),
    ("P p", W.P, W.P, placement._PARAGRAPH_BLOCK),
    ("P text", W.P, lambda: t("x"), 'text goes in a run: `r("...")`'),
    ("hyperlink tbl", W.PHyperlink, W.Tbl, placement._TABLE_SIBLING),
    ("hyperlink p", W.PHyperlink, W.P, placement._PARAGRAPH_BLOCK),
    ("hyperlink text", W.PHyperlink, lambda: t("x"), 'text goes in a run: `r("...")`'),
    # R: an R, a P or a Tbl
    ("R r", W.R, lambda: r("x"), "a run can't hold a run: put it in the paragraph, beside this one"),
    ("R p", W.R, W.P, placement._PARAGRAPH_BLOCK),
    ("R tbl", W.R, W.Tbl, placement._TABLE_SIBLING),
    # SdtContent in SdtContent
    ("sdtContent in block sdtContent", W.SdtContentBlock, W.SdtContentBlock, "a `w:sdtContent` is a content control's own part"),
    ("sdtContent in run sdtContent", W.CTSdtContentRun, W.CTSdtContentRun, "a `w:sdtContent` is a content control's own part"),
]


@pytest.mark.parametrize("case,holder,item,hint", JAVA, ids=[case for case, *_ in JAVA])
def test_java_check_parent_s_cases_are_refused_here_too(case, holder, item, hint):
    error = refused(holder(), "content", item())
    assert error.hint.startswith(hint)


#: docx4j CR-064 section 2's nine rows, and its phase 1 sets (section 8): (row, holder, field, item, hint start).
CR064 = [
    ("1 run in body", W.Body, "content", lambda: r("x"), "a run goes in a paragraph"),
    ("1 text in tc", W.Tc, "content", lambda: t("x"), "text goes in a run, in a paragraph"),
    ("1 run in block sdt content", W.SdtContentBlock, "content", lambda: r("x"), "a run goes in a paragraph"),
    ("1 run in a header", W.Hdr, "content", lambda: r("x"), "a run goes in a paragraph"),
    ("1 run in a footnote", W.CTFtnEdn, "content", lambda: r("x"), "a run goes in a paragraph"),
    ("1 run in a comment", W.CommentsComment, "content", lambda: r("x"), "a run goes in a paragraph"),
    ("1 run in a text box", W.CTTxbxContent, "content", lambda: r("x"), "a run goes in a paragraph"),
    ("1 br in body", W.Body, "content", el.br, "a break goes in a run, in a paragraph"),
    ("1 drawing in tc", W.Tc, "content", el.drawing, "a drawing goes in a run, in a paragraph"),
    ("2 drawing in p", W.P, "content", el.drawing, "a drawing goes in a run"),
    ("3 tbl in p", W.P, "content", W.Tbl, "a table is the paragraph's sibling"),
    ("3 text in p", W.P, "content", lambda: t("x"), "text goes in a run"),
    ("4 string in p", W.P, "content", lambda: "x", "a string is not an element"),
    ("4 string in r", W.R, "content", lambda: "x", "a string is not an element"),
    ("4 string in body", W.Body, "content", lambda: "x", "a string is not an element"),
    ("4 string in tbl", W.Tbl, "content", lambda: "x", "a string is not an element: text goes in a cell"),
    ("5 sectPr in body", W.Body, "content", el.sectPr, "`sectPr` is a property, not content: set `body.sect_pr`"),
    ("5 pPr in p", W.P, "content", el.pPr, "`pPr` is a property, not content: set `p.p_pr`"),
    ("5 rPr in r", W.R, "content", el.rPr, "`rPr` is a property, not content: set `r.r_pr`"),
    ("5 tblPr in tbl", W.Tbl, "content", el.tblPr, "`tblPr` is a property, not content: set `tbl.tbl_pr`"),
    ("5 trPr in tr", W.Tr, "content", el.trPr, "`trPr` is a property, not content: set `tr.tr_pr`"),
    ("5 tcPr in tc", W.Tc, "content", el.tcPr, "`tcPr` is a property, not content: set `tc.tc_pr`"),
    ("6 p in tr", W.Tr, "content", lambda: p("x"), "a paragraph goes in a cell"),
    ("6 r in tr", W.Tr, "content", lambda: r("x"), "a run goes in a paragraph, in a cell"),
    ("6 tbl in tr", W.Tr, "content", W.Tbl, "a nested table goes in a table cell"),
    ("6 tc in tbl", W.Tbl, "content", lambda: tc("x"), "a table holds rows"),
    ("6 p in tbl", W.Tbl, "content", lambda: p("x"), "a paragraph goes in a cell, in a row"),
    ("6 r in tbl", W.Tbl, "content", lambda: r("x"), "a run goes in a paragraph, in a cell"),
    ("6 tbl in tbl", W.Tbl, "content", W.Tbl, "a nested table goes in a table cell"),
    ("6 tc in body", W.Body, "content", lambda: tc("x"), "a cell goes in a row, in a table"),
    ("6 tr in body", W.Body, "content", lambda: tr(["x"]), "a row goes in a table"),
    ("7 br in p", W.P, "content", el.br, "a break goes in a run: `r(br())`"),
    ("7 tab in p", W.P, "content", el.tab, "a tab goes in a run: `r(tab())`"),
    ("8 tbl in fldSimple", W.CTSimpleField, "content", W.Tbl, "a table is the paragraph's sibling"),
    ("8 p in ins", W.RunIns, "content", W.P, "a paragraph is a block"),
    ("8 text in del", W.RunDel, "content", lambda: t("x"), "text goes in a run"),
    ("8 tbl in smartTag", W.CTSmartTagRun, "content", W.Tbl, "a table is the paragraph's sibling"),
    ("8 p in run customXml", W.CTCustomXmlRun, "content", W.P, "a paragraph is a block"),
    ("8 text in run sdt content", W.CTSdtContentRun, "content", lambda: t("x"), "text goes in a run"),
    ("9 p in run sdt content", W.CTSdtContentRun, "content", W.P, "a paragraph is a block"),
    ("9 r in block sdt content", W.SdtContentBlock, "content", lambda: r("x"), "a run goes in a paragraph"),
]


@pytest.mark.parametrize("row,holder,field,item,hint", CR064, ids=[row for row, *_ in CR064])
def test_docx4j_cr064_s_mistakes_are_refused_and_named_alike(row, holder, field, item, hint):
    error = refused(holder(), field, item())
    assert error.hint.startswith(hint), error.hint


def test_cr064_s_accepted_children_go_in():
    """CR-064 phase 1's own accepted set: an R in a hyperlink, a Tr in a Tbl, a run's text, mc:AlternateContent."""
    from docx4j_py.mce import AlternateContent

    W.PHyperlink().content.append(r("x"))
    W.Tbl().content.append(tr(["x"]))
    W.R().content.append(t("x"))
    for holder in (W.P(), W.Body(), W.Tc(), W.Tbl(), W.Tr(), W.R()):
        holder.content.append(AlternateContent())
    W.Body().content.extend([el.bookmarkStart(id=1, name="b"), el.bookmarkEnd(id=1), el.altChunk()])


# ---------------------------------------------------------------------------
# the form of a content control (D6's form mismatch), and test (b)
# ---------------------------------------------------------------------------


def test_a_block_level_control_in_a_paragraph_is_refused_with_the_run_form():
    error = refused(W.P(), "content", W.SdtBlock())
    assert error.message == "`P.content` can't hold a `SdtBlock` (`w:sdt`)"
    assert error.hint == (
        'this is a block-level content control (`SdtBlock`); in a paragraph use the run form, '
        '`el.sdt_run(...)` or `sdt(..., form="run")`'
    )


def test_a_run_level_control_in_the_body_is_refused_with_the_block_form():
    error = refused(W.Body(), "content", W.SdtRun())
    assert error.hint == (
        'this is a run-level content control (`SdtRun`); at block level use the block form, '
        '`el.sdt(...)` or `sdt(..., form="block")`'
    )


def test_a_control_of_the_wrong_form_among_rows_and_cells():
    assert "among a table's rows use the row form, `el.sdt_row(...)`" in refused(W.Tbl(), "content", W.SdtBlock()).hint
    assert "among a row's cells use the cell form, `el.sdt_cell(...)`" in refused(W.Tr(), "content", W.SdtRun()).hint


def test_a_custom_xml_element_of_the_wrong_form():
    error = refused(W.P(), "content", W.CTCustomXmlBlock())
    assert error.hint == (
        "this is a block-level custom XML element (`CTCustomXmlBlock`); in a paragraph use the run form, "
        "`el.custom_xml_run(...)`"
    )


def test_the_sdt_builder_refuses_a_form_its_content_cannot_take():
    """CR-020 section 2.4 (b): ``sdt([p("a")], form="run")`` built a paragraph inside a run-level control."""
    with pytest.raises(MisplacedContentError) as raised:
        sdt([p("a")], form="run")
    assert raised.value.message == "`CTSdtContentRun.content` can't hold a `P` (`w:p`)"
    assert raised.value.hint == placement._PARAGRAPH_BLOCK
    assert isinstance(sdt([r("a")], form="run"), W.SdtRun)


# ---------------------------------------------------------------------------
# the derived hints
# ---------------------------------------------------------------------------


def test_a_derived_hint_names_where_the_element_goes():
    error = refused(W.R(), "content", el.bookmarkStart(id=1, name="b"))
    assert error.hint.startswith("`w:bookmarkStart` goes in: `Body.content`, `Tc.content`, `P.content`")
    assert error.hint.endswith(", ...")
    assert error.hint.count("`") == 2 * 6


def test_a_derived_hint_for_a_property_held_elsewhere():
    error = refused(W.P(), "content", el.rPr())
    assert error.hint.startswith("`rPr` is a property, not content: set it as `R.r_pr`")


def test_an_element_in_another_form_names_the_class_the_field_takes():
    """A math run's ``w:ins`` is ``Ins1``, not the ``RunIns`` every other holder takes (section 2.3)."""
    from docx4j_py.math import CTR

    assert placement.entry(CTR, "content").by_name[f"{{{W_NS}}}ins"] == (W.Ins1,)
    error = refused(CTR(), "content", el.ins())
    assert error.hint == "`CTR.content` takes `w:ins` in another form, an `Ins1`; build that one, `el.ins_math(...)`"


def test_an_item_with_no_element_name_lists_what_the_field_takes():
    error = refused(W.P(), "content", W.CTString())
    assert error.hint.startswith("`P.content` takes `")


def test_any_element_goes_only_where_there_is_a_wildcard():
    from docx4j_xsdata.formats.dataclass.models.generics import AnyElement

    error = refused(W.P(), "content", AnyElement(qname=f"{{{W_NS}}}r"))
    assert error.message == "`P.content` can't hold an `AnyElement` (`w:r`)"
    # a wildcard list is not checked: a branch's content takes it
    choice = import_module("docx4j_py.mce").AlternateContentChoice()
    choice.any_element.append(AnyElement(qname="{urn:x}y"))
    assert placement.entry(type(choice), "any_element") is None


def test_none_is_not_an_element():
    error = refused(W.P(), "content", None)
    assert error.message == "`P.content` can't hold a `NoneType`"


# ---------------------------------------------------------------------------
# the table (D1, D10), and the census section 2 measured
# ---------------------------------------------------------------------------


def test_the_census_is_the_model_s():
    """CR-020 sections 2.2 and 2.3, re-derived from the model (section 10.2)."""
    counts = placement.census()
    assert counts["classes"] == 1969
    assert counts["unbuildable"] == 1  # inkml's MappingType: its annotations name Math, which does not resolve
    assert counts["pairs"] == 537
    assert counts["checked"] == 509
    assert counts["unchecked"] == 28
    assert counts["compound"] + counts["element"] == 509
    assert counts["mixed"] == 6
    assert counts["names"] == 71
    assert counts["one_shape"] == 60
    assert counts["differing"] == 11


def test_the_eleven_names_whose_classes_differ():
    local = sorted(name.rpartition("}")[2] + "@" + name[1:].partition("}")[0].rsplit("/", 1)[-1] for name in placement.differing_names())
    assert local == sorted(
        [
            "sdt@main",
            "customXml@main",
            "ins@main",
            "del@main",
            "name@main",
            "alphaMod@main",
            "lum@main",
            "tint@main",
            "xfrm@main",
            "effect@main",
            "lvl@chartex",
        ]
    )


def test_a_wildcard_or_choiceless_field_has_no_table():
    unchecked_pairs = [
        (cls, field)
        for cls in placement.model_classes()
        for field in _list_field_names(cls)
        if placement.entry(cls, field) is None
    ]
    assert len(unchecked_pairs) == 28 + len(_list_field_names(import_module("docx4j_py.inkml").MappingType))
    for cls, field in unchecked_pairs:
        cls_meta = _meta_or_none(cls)
        if cls_meta is None:
            continue
        var = next(v for v in cls_meta.get_element_vars() if v.name == field)
        assert var.is_wildcard or var.wildcards or not (var.elements or var.is_element)


def _meta_or_none(cls):
    try:
        return context().build(cls)
    except Exception:  # noqa: BLE001
        return None


def test_every_namespace_is_checked():
    """D10: the table is derived, so DrawingML, VML, MathML and the rest are checked too."""
    by_module: dict[str, tuple[type, str]] = {}
    for cls, field in _checked_pairs():
        by_module.setdefault(cls.__module__, (cls, field))
    assert {"docx4j_py.wml", "docx4j_py.dml.main", "docx4j_py.dml.chart", "docx4j_py.vml.main", "docx4j_py.math"} <= set(by_module)
    for module, (cls, field) in sorted(by_module.items()):
        if module == "docx4j_py.wml":
            continue
        error = refused(cls(), field, W.Tbl())
        assert error.code == "target.misplaced", module
    assert len(by_module) > 30, sorted(by_module)


def _checked_pairs():
    return [
        (cls, field)
        for cls in placement.model_classes()
        for field in _list_field_names(cls)
        if placement.entry(cls, field) is not None
    ]


def test_the_table_is_built_lazily():
    """Nothing is added at import: a fresh process has no table until a list is checked."""
    import subprocess
    import sys

    code = (
        "import docx4j_py.wml as W, docx4j_py.placement as pl\n"
        "assert not pl._entries, pl._entries\n"
        "W.P().content.append(W.R())\n"
        "assert set(pl._entries) == {(W.P, 'content')}, pl._entries\n"
        "print('ok')\n"
    )
    out = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, cwd=str(ROOT), check=True)
    assert out.stdout.strip() == "ok"


# ---------------------------------------------------------------------------
# the legal placements, exhaustively, from el
# ---------------------------------------------------------------------------


def _el_factories() -> dict[type, object]:
    """Each class some ``el`` factory builds, with the factory: the instances come from ``el``, not the table."""
    built: dict[type, object] = {}
    for path in sorted(set(EL_MODULES.values())):
        module = import_module(path)
        for name in module.__all__:
            factory = getattr(module, name)
            if not callable(factory) or isinstance(factory, type):
                continue
            try:
                instance = factory()
            except Exception:  # noqa: BLE001, S112 - a factory that wants an argument
                continue
            built.setdefault(type(instance), factory)
    return built


def _choices(cls: type, field: str) -> set[type]:
    """The field's element classes, read from the metadata directly (not from the table)."""
    meta = context().build(cls)
    var = next(v for v in meta.get_element_vars() if v.name == field)
    alternatives = var.elements.values() if var.elements else (var,)
    return {tp for alt in alternatives for tp in alt.types if dataclasses.is_dataclass(tp)}


#: Every choice class of every checkable field, one placement each (CR-020 section 10.8): built by an
#: ``el`` factory, and by the class itself for the 77 classes no factory builds (VML's ``x:ClientData``
#: children, inkml's types, chartex's intermediate classes), 155 placements.
FROM_EL = 3480
BY_CLASS = 155


def test_every_legal_placement_goes_in():
    """For every checkable (class, field), an instance of each choice class goes in: the instances from ``el``."""
    factories = _el_factories()
    from_el = by_class = 0
    for cls, field in _checked_pairs():
        holder = cls()
        items = getattr(holder, field)
        everything = []
        for choice in sorted(_choices(cls, field), key=lambda c: (c.__module__, c.__name__)):
            factory = factories.get(choice)
            if factory is None:
                item = choice()
                by_class += 1
            else:
                item = factory()
                from_el += 1
            items.append(item)
            everything.append(type(item))
        # and all at once, through the constructor and through extend
        cls(**{field: [kind() for kind in everything]})
        cls().__getattribute__(field).extend(kind() for kind in everything)
    assert (from_el, by_class) == (FROM_EL, BY_CLASS)


# ---------------------------------------------------------------------------
# the shape rule (D2 rule 2)
# ---------------------------------------------------------------------------


def test_the_hyperlink_case_is_accepted_and_saved_as_a_hyperlink():
    """CR-020 section 2.4 (a): a ``PHyperlink`` in a run-level control's content, whose class there differs."""
    content = W.CTSdtContentRun()
    link = el.hyperlink(content=[r("x")])
    assert type(link) is W.PHyperlink
    assert W.PHyperlink not in placement.entry(W.CTSdtContentRun, "content").classes
    content.content.append(link)
    xml = to_xml(W.SdtRun(sdt_content=content), name=f"{{{W_NS}}}sdt")
    root = etree.fromstring(xml.encode())
    assert [etree.QName(child).localname for child in root.find(f"{{{W_NS}}}sdtContent")] == ["hyperlink"]


def test_the_content_api_s_hyperlink_in_a_control_goes_through_the_check():
    """``hyperlinks._wrap`` puts a ``PHyperlink`` in ``CTSdtContentRun.content``: accepted by rule 2."""
    import io
    import zipfile

    from docx4j_py import create_package

    pkg = create_package()
    paragraph = pkg.body.insert_paragraph("see this page now")
    paragraph.search("this page")[0].insert_content_control("RichText")
    paragraph.search("this")[0].hyperlink = "https://example.com/inside"
    control = next(item for item in paragraph.element.content if isinstance(item, W.SdtRun))
    links = [item for item in control.sdt_content.content if isinstance(item, W.PHyperlink)]
    assert len(links) == 1
    assert links[0].parent is control.sdt_content
    with zipfile.ZipFile(io.BytesIO(pkg.save())) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
    content = root.find(f".//{{{W_NS}}}sdt/{{{W_NS}}}sdtContent")
    assert [etree.QName(child).localname for child in content] == ["hyperlink", "r"]
    assert content[0].findtext(f".//{{{W_NS}}}t") == "this"


def test_a_block_level_control_in_a_paragraph_is_refused_by_shape():
    """By element name alone it would go in: ``P.content`` takes ``w:sdt``, as ``SdtRun``."""
    assert f"{{{W_NS}}}sdt" in placement.entry(W.P, "content").by_name
    assert placement.shape_of(W.SdtBlock) != placement.shape_of(W.SdtRun)
    refused(W.P(), "content", W.SdtBlock())


def _multi_class_names() -> dict[str, set[type]]:
    names: dict[str, set[type]] = {}
    for cls in placement.model_classes():
        meta = _meta_or_none(cls)
        if meta is None:
            continue
        for var in meta.get_element_vars():
            for qname, alt in (var.elements or {}).items():
                names.setdefault(qname, set()).update(tp for tp in alt.types if dataclasses.is_dataclass(tp))
    return {name: classes for name, classes in names.items() if len(classes) > 1}


def test_each_differing_name_is_refused_in_the_holder_of_another_of_its_classes():
    multi = _multi_class_names()
    tried = 0
    for name in placement.differing_names():
        for cls, field in placement.holders_of(name):
            found = placement.entry(cls, field)
            if found is None or not found.compound:
                continue
            mine = set(found.by_name[name])
            for other in multi[name] - mine:
                if any(placement.shape_of(other) == placement.shape_of(c) for c in mine):
                    continue
                with pytest.raises(MisplacedContentError):
                    getattr(cls(), field).append(other())
                tried += 1
    assert tried > 30, tried


def test_an_accepted_stranger_is_saved_as_the_element_the_field_takes():
    """Rule 2's third condition against the serialiser itself: every class of a one-shape name, in every holder."""
    multi = _multi_class_names()
    accepted = refused_count = 0
    ser = serializer()
    for cls, field in _checked_pairs():
        found = placement.entry(cls, field)
        if not found.compound:
            continue
        for name, mine in found.by_name.items():
            for other in multi.get(name, ()):
                if other in mine:
                    continue
                with unchecked():
                    holder = cls(**{field: [other()]})
                root = etree.fromstring(ser.render(holder).encode())
                written = [child.tag for child in root]
                xsi = [child.get("{http://www.w3.org/2001/XMLSchema-instance}type") for child in root]
                if placement.verdict(cls, field, other()):
                    assert written == [name] and xsi == [None], (cls, field, other, written)
                    accepted += 1
                else:
                    refused_count += 1
    assert accepted > 250, accepted
    assert refused_count > 600, refused_count


# ---------------------------------------------------------------------------
# where it runs (D3): every mutation, and a refused one changes nothing
# ---------------------------------------------------------------------------


def test_every_mutation_is_checked():
    paragraph = p("a")
    first = paragraph.content[0]
    good = r("b")
    for mutate in (
        lambda items: items.append(W.Tbl()),
        lambda items: items.insert(0, W.Tbl()),
        lambda items: items.extend([r("c"), W.Tbl()]),
        lambda items: items.__iadd__([r("c"), W.Tbl()]),
        lambda items: items.__setitem__(0, W.Tbl()),
        lambda items: items.__setitem__(slice(0, 1), [good, W.Tbl()]),
    ):
        with pytest.raises(MisplacedContentError):
            mutate(paragraph.content)
        assert list(paragraph.content) == [first]
        assert first.parent is paragraph
        assert good.parent is None


def test_a_refused_extend_or_slice_assignment_leaves_the_list_unchanged():
    paragraph = p("a", "b")
    before = list(paragraph.content)
    incoming = [r("c"), r("d"), W.Tbl()]
    with pytest.raises(MisplacedContentError):
        paragraph.content.extend(incoming)
    assert list(paragraph.content) == before
    with pytest.raises(MisplacedContentError):
        paragraph.content[0:2] = incoming
    assert list(paragraph.content) == before
    assert all(item.parent is paragraph for item in before)
    assert all(item.parent is None for item in incoming[:2])
    with pytest.raises(MisplacedContentError):
        paragraph.content += iter(incoming)
    assert list(paragraph.content) == before


def test_the_constructor_is_checked():
    with pytest.raises(MisplacedContentError):
        W.P(content=[W.Tbl()])
    with pytest.raises(MisplacedContentError):
        W.P(content=ChildList([W.Tbl()]))
    with pytest.raises(MisplacedContentError):
        el.p(content=[W.Tbl()])
    with pytest.raises(MisplacedContentError):
        p(W.Tbl())
    with pytest.raises(MisplacedContentError):
        tc(r("x"))


def test_a_list_with_a_field_name_is_checked_and_one_without_is_not():
    run = r("x")
    ChildList([W.Tbl()], owner=run)  # no field name: unchecked, as an unowned list is unparented
    ChildList([W.Tbl()])
    with pytest.raises(MisplacedContentError):
        ChildList([W.Tbl()], owner=run, field="content")
    items = ChildList(owner=run, field="content")
    assert items.field == "content"
    with pytest.raises(MisplacedContentError):
        items.append(W.Tbl())


def test_set_owner_and_link_parents_check_a_list_handed_over_whole():
    run = W.R()
    run.content = ChildList([W.Tbl()])  # a whole-list assignment is the validator's (Q4) ...
    with pytest.raises(MisplacedContentError):
        link_parents(run)  # ... until the list is bound to its owner
    items = ChildList([W.Tbl()])
    with pytest.raises(MisplacedContentError):
        items.set_owner(W.R(), "content")
    assert items.owner is None and items.field is None


def test_a_list_re_bound_to_its_own_owner_is_not_checked_again():
    run = r("x")
    with unchecked():
        run.content.append(W.Tbl())
    link_parents(run)  # the same owner and field: nothing new went in
    run.content.set_owner(run, "content")
    assert isinstance(run.content[-1], W.Tbl)


# ---------------------------------------------------------------------------
# unchecked() (D9) and the parse (D4)
# ---------------------------------------------------------------------------


def test_unchecked_lets_an_invalid_tree_be_built_on_purpose():
    with unchecked():
        run = r("x")
        run.content.append(W.Tbl())
        paragraph = W.P(content=[W.Tbl(), "loose text"])
    assert isinstance(run.content[-1], W.Tbl)
    assert len(paragraph.content) == 2
    with pytest.raises(MisplacedContentError):
        run.content.append(W.Tbl())


def test_unchecked_nests_and_restores():
    with unchecked():
        with unchecked():
            W.R().content.append(W.Tbl())
        W.R().content.append(W.Tbl())
    with pytest.raises(MisplacedContentError):
        W.R().content.append(W.Tbl())


def test_unchecked_restores_the_check_after_an_exception():
    with pytest.raises(RuntimeError), unchecked():
        raise RuntimeError("boom")
    with pytest.raises(MisplacedContentError):
        W.R().content.append(W.Tbl())


def test_a_deep_copy_of_a_tree_built_unchecked_needs_no_unchecked():
    with unchecked():
        run = r("x")
        run.content.append(W.Tbl())
    copied = deep_copy(run)
    assert isinstance(copied.content[-1], W.Tbl)
    assert copied.content[-1].parent is copied
    with pytest.raises(MisplacedContentError):
        copied.content.append(W.Tbl())


def test_unchecked_is_per_thread():
    started = threading.Event()
    release = threading.Event()
    results: list[object] = []

    def other() -> None:
        with unchecked():
            started.set()
            release.wait(5)

    thread = threading.Thread(target=other)
    thread.start()
    started.wait(5)
    try:
        W.R().content.append(W.Tbl())
    except MisplacedContentError as exc:
        results.append(exc)
    release.set()
    thread.join()
    assert len(results) == 1


def test_a_parse_is_not_checked(monkeypatch):
    """D4: the parser's ``class_factory`` builds with the check off; nothing is judged during a parse."""
    from docx4j_py import child

    calls = []
    real = child._check

    def counting(owner, field, items):
        calls.append((type(owner).__name__, field))
        return real(owner, field, items)

    run = r("c")
    monkeypatch.setattr(child, "_check", counting)
    parsed = wml(f'<w:p xmlns:w="{W_NS}"><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>')
    assert calls == []
    assert len(parsed.content) == 2
    parsed.content.append(run)
    assert calls == [("P", "content")]


def test_the_check_is_on_again_after_a_parse_and_after_one_that_raises():
    wml(f'<w:p xmlns:w="{W_NS}"><w:r><w:t>a</w:t></w:r></w:p>')
    with pytest.raises(MisplacedContentError):
        W.R().content.append(W.Tbl())
    with pytest.raises(Exception) as raised:
        wml(f'<w:p xmlns:w="{W_NS}"><w:r><w:bogus/></w:r></w:p>')  # strict: an unknown element raises
    assert not isinstance(raised.value, MisplacedContentError)
    with pytest.raises(MisplacedContentError):
        W.R().content.append(W.Tbl())


def test_every_sample_loads_with_no_check(monkeypatch):
    """The corpus parses with no refusal and no check: every part of every sample, unmarshalled."""
    from docx4j_py import child
    from docx4j_py.openpackaging import OpcPackage

    calls = []
    monkeypatch.setattr(child, "_check", lambda owner, field, items: calls.append(field))
    assert len(SAMPLES) == 16
    for path in SAMPLES:
        with OpcPackage.load(path) as pkg:
            pkg.unmarshal_all()
            assert pkg.skipped == []
    assert calls == []


def test_the_engine_s_parser_parses_unchecked():
    """Q3 (section 10.5): a context variable set around each parse, in the one parser the package builds."""
    from docx4j_py.runtime import UncheckedXmlParser, parser

    assert isinstance(parser(), UncheckedXmlParser)


def test_every_parser_in_the_package_is_runtime_parser_s():
    """So that every parse in the package is unchecked: no module but ``runtime.parser`` builds one."""
    import docx4j_py

    package = Path(docx4j_py.__file__).parent
    built = []
    for path in sorted(package.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                for inner in ast.walk(node):
                    if isinstance(inner, ast.Call):
                        name = getattr(inner.func, "id", getattr(inner.func, "attr", None))
                        if name in ("XmlParser", "TreeParser", "UncheckedXmlParser", "UserXmlParser"):
                            built.append((path.relative_to(package).as_posix(), node.name, name))
    assert sorted(set(built)) == [("runtime.py", "parser", "UncheckedXmlParser")]


# ---------------------------------------------------------------------------
# the engine's own sites (D8)
# ---------------------------------------------------------------------------


def test_every_owned_child_list_the_engine_builds_names_its_field():
    """A ``ChildList(..., owner=x)`` with no ``field=`` would be unchecked; every one in the package names it."""
    import docx4j_py

    package = Path(docx4j_py.__file__).parent
    offenders = []
    for path in sorted(package.rglob("*.py")):
        relative = path.relative_to(package).as_posix()
        if relative == "child.py":
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.Call) and getattr(node.func, "id", getattr(node.func, "attr", None)) == "ChildList":
                keywords = {keyword.arg for keyword in node.keywords}
                if "owner" in keywords and "field" not in keywords:
                    offenders.append(f"{relative}:{node.lineno}")
    assert offenders == []


def test_every_field_name_the_engine_gives_a_list_is_a_list_field():
    """A misspelt ``field=`` would leave the list unchecked, silently: each literal names some class's list field."""
    import docx4j_py

    known = {field for cls in placement.model_classes() for field in _list_field_names(cls)}
    package = Path(docx4j_py.__file__).parent
    given = []
    for path in sorted(package.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.Call) and getattr(node.func, "id", getattr(node.func, "attr", None)) == "ChildList":
                for keyword in node.keywords:
                    if keyword.arg == "field" and isinstance(keyword.value, ast.Constant):
                        given.append((path.name, node.lineno, keyword.value.value))
    assert len(given) >= 12
    assert [entry for entry in given if entry[2] not in known] == []
