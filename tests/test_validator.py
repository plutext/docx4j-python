"""The validator, and the save's check. CR-020 Phase B (sections 4.7 and 11).

``misplaced(root)`` (:mod:`docx4j_py.validator`, :func:`docx4j_py.traversal.misplaced`,
``package.misplaced()``) walks a tree, a part or a package and returns every finding:

* ``target.misplaced``: what Phase A's add-time check refuses (D2), wherever it got past
  that check (under ``unchecked()``, a plain list assigned whole, a single field assigned,
  ``object.__setattr__``);
* ``cell.no_end_paragraph``: a table cell that does not end with a paragraph (CR-016);
* ``value.out_of_range``: an integer measurement past what Word stores (CR-022's table,
  read side), held to Word 365's saves of docx4j CR-066 rounds 5 and 6.

``save(check=True)`` raises for what the session introduced and records what the source
held.

    .venv-fork/bin/python -m pytest tests/test_validator.py -q
"""

from __future__ import annotations

import io
import os
import re
import zipfile
from functools import cache
from pathlib import Path

import pytest
from lxml import etree
from test_misplaced import CR064, JAVA, SIX, _checked_pairs, _multi_class_names

import docx4j_py.wml as W
from docx4j_py import placement
from docx4j_py.child import ChildList, unchecked
from docx4j_py.model.content import SaveCheckError
from docx4j_py.model.values import integer, integer_limit, integer_limits
from docx4j_py.model.word_integers import (
    STORAGE,
    Storage,
    kept_range,
    read_range,
    reading,
    storage_of,
)
from docx4j_py.openpackaging.api import create_package, load
from docx4j_py.openpackaging.save import checking, saves_check
from docx4j_py.traversal import misplaced
from docx4j_py.validator import (
    CellWithoutEndParagraph,
    Finding,
    MisplacedContent,
    ValueOutOfRange,
    findings_in,
)
from docx4j_py.wml import el, p, r, tbl, tr

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "tests" / "fixtures"
LIMITS = FIXTURES / "integer-limits"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

# the probes hold values the converters refuse ("720.5", a bad colour): each is kept and warned about
pytestmark = pytest.mark.filterwarnings("ignore::docx4j_xsdata.exceptions.ConverterWarning")


def _only(found: list[Finding], kind: type) -> list:
    return [f for f in found if isinstance(f, kind)]


# ---------------------------------------------------------------------------
# the findings
# ---------------------------------------------------------------------------


def test_a_finding_is_plain_data_with_to_dict():
    with unchecked():
        run = r("x")
        run.content.append(W.Tbl())
    (finding,) = findings_in(p(run))
    assert isinstance(finding, MisplacedContent)
    assert finding.code == "target.misplaced"
    assert finding.to_dict() == {
        "code": "target.misplaced",
        "path": "/w:p/w:r[1]/w:tbl[1]",
        "holder": "R",
        "field": "content",
        "index": 1,
        "item": "Tbl",
        "element": "w:tbl",
        "message": "`R.content` can't hold a `Tbl` (`w:tbl`)",
        "hint": placement._TABLE_SIBLING,
    }
    assert "node" not in repr(finding)
    with pytest.raises(AttributeError):
        finding.code = "x"  # type: ignore[misc]


def test_the_message_and_hint_are_the_add_time_error_s():
    for _case, holder, field, item, _hint in SIX:
        error = placement.refusal(holder(), field, item())
        made = holder()
        with unchecked():
            getattr(made, field).append(item())
        (finding,) = _only(findings_in(made), MisplacedContent)
        assert (finding.message, finding.hint) == (error.message, error.hint)


def test_the_codes():
    assert MisplacedContent.code == "target.misplaced"
    assert CellWithoutEndParagraph.code == "cell.no_end_paragraph"
    assert ValueOutOfRange.code == "value.out_of_range"


# ---------------------------------------------------------------------------
# misplaced content: Phase A's refused cases, wherever they got past the check
# ---------------------------------------------------------------------------

#: Phase A's tables: (holder, field, item factory) for every case its add-time check refuses.
REFUSED = [
    *((holder, field, item) for _c, holder, field, item, _h in SIX),
    *((holder, "content", item) for _c, holder, item, _h in JAVA),
    *((holder, field, item) for _c, holder, field, item, _h in CR064),
    (W.P, "content", W.SdtBlock),
    (W.Body, "content", W.SdtRun),
    (W.Tbl, "content", W.SdtBlock),
    (W.Tr, "content", W.SdtRun),
    (W.P, "content", W.CTCustomXmlBlock),
    (W.R, "content", lambda: el.bookmarkStart(id=1, name="b")),
    (W.P, "content", el.rPr),
    (W.P, "content", W.CTString),
    (W.P, "content", lambda: None),
]


def _ids(cases):
    return [f"{holder.__name__}.{field}-{index}" for index, (holder, field, _item) in enumerate(cases)]


@pytest.mark.parametrize(("holder", "field", "item"), REFUSED, ids=_ids(REFUSED))
def test_phase_a_s_refused_cases_are_found_built_under_unchecked(holder, field, item):
    made = holder()
    value = item()
    with unchecked():
        getattr(made, field).append(value)
    (finding,) = _only(findings_in(made), MisplacedContent)
    assert (finding.holder, finding.field, finding.index) == (holder.__name__, field, 0)
    assert finding.node is value
    assert finding.message == placement.refusal(made, field, value).message


@pytest.mark.parametrize(("holder", "field", "item"), REFUSED, ids=_ids(REFUSED))
def test_phase_a_s_refused_cases_are_found_in_a_plain_list_assigned_whole(holder, field, item):
    """``holder.content = [x]``: no ``ChildList``, so no add-time check (CR-020 section 10.11 item 4)."""
    made = holder()
    setattr(made, field, [item()])
    assert type(getattr(made, field)) is list
    (finding,) = _only(findings_in(made), MisplacedContent)
    assert finding.index == 0


@pytest.mark.parametrize(("holder", "field", "item"), REFUSED, ids=_ids(REFUSED))
def test_phase_a_s_refused_cases_are_found_through_object_setattr(holder, field, item):
    made = holder()
    object.__setattr__(made, field, ChildList([item()]))
    assert len(_only(findings_in(made), MisplacedContent)) == 1


def test_a_single_field_assigned_the_wrong_class():
    """``p.p_pr = R()``: a single-valued field is never adopted, so only the validator sees it."""
    paragraph = p("x")
    paragraph.p_pr = r("y")
    (finding,) = findings_in(paragraph)
    assert (finding.holder, finding.field, finding.index, finding.item, finding.element) == ("P", "p_pr", None, "R", "w:r")
    assert finding.message == "`P.p_pr` can't hold an `R` (`w:r`)"
    assert finding.path == "/w:p/w:r[1]"
    run = r("x")
    object.__setattr__(run, "r_pr", el.pPr())
    (finding,) = findings_in(run)
    assert (finding.field, finding.item, finding.element) == ("r_pr", "PPr", "w:pPr")


def test_a_single_field_s_legal_values_are_not_found():
    paragraph = p("x")
    paragraph.p_pr = el.pPr()
    run = paragraph.content[0]
    run.r_pr = el.rPr(b=el.b())
    assert findings_in(paragraph) == []


def test_a_subclass_written_with_an_xsi_type_is_found():
    """Word writes no ``xsi:type``: a ``PPr`` in ``w:pPrChange``, declared ``CT_PPrBase`` (``deep_copy_as``'s case)."""
    change = el.pPrChange(id=1, author="a")
    change.p_pr = W.PPr()
    (finding,) = findings_in(change)
    assert (finding.field, finding.item) == ("p_pr", "PPr")
    change.p_pr = W.PPrBase()
    assert findings_in(change) == []


def test_every_stranger_the_check_refuses_is_found_and_every_one_it_accepts_is_not():
    """Section 10.3's 1,440 placements, each built under ``unchecked()``: the validator agrees with D2."""
    multi = _multi_class_names()
    found_count = clean = 0
    for cls, field in _checked_pairs():
        entry = placement.entry(cls, field)
        if not entry.compound:
            continue
        for name, mine in entry.by_name.items():
            for other in multi.get(name, ()):
                if other in mine:
                    continue
                with unchecked():
                    holder = cls(**{field: [other()]})
                found = [f for f in _only(findings_in(holder), MisplacedContent) if f.holder == cls.__name__]
                if placement.verdict(cls, field, other()):
                    assert not found, (cls, field, other)
                    clean += 1
                else:
                    assert len(found) == 1 and found[0].index == 0, (cls, field, other)
                    found_count += 1
    assert (clean, found_count) == (300, 1140)


def test_a_finding_deep_in_a_tree_has_its_path():
    body = W.Body(content=[p("a"), tbl([["x", "y"]]), p("b")])
    cell = body.content[1].content[0].content[1]
    with unchecked():
        cell.content[0].content.append(W.Tbl())
    (finding,) = findings_in(body)
    assert finding.path == "/w:body/w:tbl[1]/w:tr[1]/w:tc[2]/w:p[1]/w:tbl[1]"
    assert finding.holder == "P"


def test_a_string_s_path_is_its_holder_s():
    with unchecked():
        paragraph = p("a")
        paragraph.content.append("text")
    (finding,) = findings_in(W.Body(content=[paragraph]))
    assert (finding.path, finding.item, finding.element) == ("/w:body/w:p[1]", "str", None)


def test_a_misplaced_item_is_walked_into():
    """A table in a run is reported, and so is what is wrong inside it."""
    inner = W.Tbl(content=[tr(["x"])])
    with unchecked():
        inner.content[0].content[0].content.append(r("run in a cell"))
        run = r("x")
        run.content.append(inner)
    found = findings_in(p(run))
    assert [(f.holder, f.item) for f in found] == [("R", "Tbl"), ("Tc", "R")]


def test_the_engine_s_own_xsi_type_in_the_core_properties_is_legal():
    """``dcterms:created`` takes any type by ``xsi:type`` (OPC's ``dcterms:W3CDTF``): bare, as the engine sets it,
    and as a ``DerivedElement``, as a parse binds it (CR-020 section 11)."""
    package = create_package()
    core = next(part for part in package.parts.parts() if str(part.part_name) == "/docProps/core.xml")
    assert type(core.contents.created).__name__ == "W3CDTF"
    assert misplaced(core) == []
    loaded = load(io.BytesIO(package.save()))
    core = next(part for part in loaded.parts.parts() if str(part.part_name) == "/docProps/core.xml")
    assert type(core.read_contents().created).__name__ == "DerivedElement"
    assert misplaced(core) == []


def test_only_an_element_declared_with_no_type_takes_any_xsi_type():
    """``_any_type``: the 48 Dublin Core elements declared with no type of their own (``dcterms:created``); a named
    type with wildcard content (``w:CT_ShapeDefaults``) takes only its own derivations, so an ``R`` there is found."""
    any_types = [cls for cls in placement.model_classes() if placement._any_type(cls)]
    assert len(any_types) == 48
    assert {cls.__module__ for cls in any_types} == {"docx4j_py.docprops.dc", "docx4j_py.docprops.dcterms"}
    assert not placement._any_type(W.CTShapeDefaults) and not placement._any_type(W.CTPictureBase)
    settings = W.Settings()
    settings.shape_defaults = r("x")
    (finding,) = findings_in(settings)
    assert (finding.field, finding.item) == ("shape_defaults", "R")


# ---------------------------------------------------------------------------
# the cell rule (CR-016)
# ---------------------------------------------------------------------------


def _table(*cells):
    return W.Tbl(content=[W.Tr(content=list(cells))])


def test_a_cell_that_ends_with_a_table_or_holds_no_block_is_found():
    nested = W.Tc(content=[p("a"), tbl([["x"]])])
    empty = W.Tc()
    marked = W.Tc(content=[p("a"), el.bookmarkStart(id=1, name="b"), el.bookmarkEnd(id=1)])
    body = W.Body(content=[_table(nested, empty, marked)])
    found = findings_in(body)
    assert [(type(f), f.path, f.ends_with) for f in found] == [
        (CellWithoutEndParagraph, "/w:body/w:tbl[1]/w:tr[1]/w:tc[1]", "w:tbl"),
        (CellWithoutEndParagraph, "/w:body/w:tbl[1]/w:tr[1]/w:tc[2]", None),
    ]
    assert found[0].node is nested
    assert found[0].to_dict()["code"] == "cell.no_end_paragraph"
    assert "ends with a table" in found[0].message and "no paragraph or table" in found[1].message


def test_a_cell_read_through_its_wrappers():
    """A paragraph in a block control ends the cell; a table there does not."""
    ok = W.Tc(content=[el.sdt(sdt_content=W.SdtContentBlock(content=[p("a")]))])
    wrapped_table = W.Tc(content=[el.sdt(sdt_content=W.SdtContentBlock(content=[tbl([["x"]])]))])
    found = findings_in(_table(ok, wrapped_table))
    assert [(f.path, f.ends_with) for f in found] == [("/w:tbl/w:tr[1]/w:tc[2]", "w:tbl")]


def test_the_engine_s_tables_end_their_cells_with_paragraphs():
    assert findings_in(W.Body(content=[tbl([["a", "b"], ["c", "d"]])])) == []


# ---------------------------------------------------------------------------
# the wrap model and the storage strings
# ---------------------------------------------------------------------------

#: The storage strings docx4j-core-ts reports (the objects package's ``storageOf``, at core-ts ``b7c2680``), by attribute
#: path of the limits table, and for a path whose reading depends on the rule, by rule.
CORE_TS_STORAGE = {
    "w:ind/@w:left": "int16",
    "w:ind/@w:firstLine": "int16, shared by firstLine (+) and hanging (-)",
    "w:ind/@w:hanging": "int16, shared by firstLine (+) and hanging (-)",
    "w:spacing/@w:before": "uint16",
    "w:pgMar/@w:left": "uint16",
    "w:tab/@w:pos": "int16, then ±31680",
    "w:pgSz/@w:w": "uint16, then 144..31681",
    "w:pgSz/@w:h": "uint16, then at least 144",
    "w:pgNumType/@w:start": "int32",
    "w:col/@w:w": "int16, a negative computed",
    "w:col/@w:space": "int16, a negative written -1",
    "w:framePr/@w:h": "15 bits, bit 15 clearing exact",
    "w:top/@w:sz": "8 bits",
    "w:top/@w:space": "5 bits",
    "w:trHeight/@w:val": "int16, the sign the rule",
    "w:spacing/@w:line": "int16, the sign the rule",
}


def test_every_row_has_a_storage_and_every_storage_a_row():
    assert {row.attribute for row in integer_limits()} == set(STORAGE)


def test_the_storage_strings_are_core_ts_s():
    for attribute, token in CORE_TS_STORAGE.items():
        assert storage_of(attribute).token == token, attribute
    # a line under auto reads as plain int16; a row height under auto is ignored
    assert reading(storage_of("w:spacing/@w:line"), "line", 70000, "auto").storage == "int16"
    assert reading(storage_of("w:spacing/@w:line"), "line", 70000, None).storage == "int16"
    assert reading(storage_of("w:trHeight/@w:val"), "val", 70000, "auto").storage.endswith("; ignored under auto")
    tokens = {s.token for s in STORAGE.values()} | {"int16"}
    assert tokens == {
        "int16", "uint16", "int32", "8 bits", "5 bits", "int16, shared by firstLine (+) and hanging (-)",
        "int16, the sign the rule", "int16, then ±31680", "uint16, then 144..31681", "uint16, then at least 144",
        "int16, a negative computed", "int16, a negative written -1", "15 bits, bit 15 clearing exact",
    }


SAMPLES = [-70000, -65537, -65536, -65535, -32769, -32768, -32767, -31681, -31680, -256, -255, -32, -31, -2, -1, 0, 1, 31,
           32, 143, 144, 255, 256, 31680, 31681, 31682, 32766, 32767, 32768, 32769, 65534, 65535, 65536, 65537, 70000,
           1205003, 7205001, 2147483647, 2147483648, -2147483649]


@pytest.mark.parametrize("row", integer_limits(), ids=lambda row: f"{row.attribute}{row.context or ''}")
def test_the_wrap_model_keeps_exactly_the_row_s_range_as_read(row):
    """``kept_range`` (core-ts's ``keptRange``) is the row's range with ``word_writes`` put back; ``reading().kept``
    is True exactly there, and at Word's own spellings (-1 for 65,535; ``w:col/@w:space``'s sentinel)."""
    storage = storage_of(row.attribute)
    local = row.name.removeprefix("w:")
    rule = row.context[1][0] if row.context and row.context[0] in ("w:lineRule", "w:hRule") else None
    assert kept_range(storage, local, rule) == read_range(row)
    low, high = read_range(row)
    spellings = {integer(spelling) for _stored, spelling in row.word_writes}
    for value in {*SAMPLES, low - 1, low, low + 1, high - 1, high, high + 1}:
        kept = reading(storage, local, value, rule).kept
        expected = low <= value <= high or value == row.sentinel or value in spellings
        assert kept == expected, (value, kept)


def test_the_readings_core_ts_states():
    """core-ts's message cases: 7205001 in ``w:left`` -3959; exact 65535 atLeast 1; ``w:firstLine`` 32768 hanging."""
    assert reading(storage_of("w:ind/@w:left"), "left", 7205001).value == -3959
    line = reading(storage_of("w:spacing/@w:line"), "line", 65535, "exact")
    assert (line.value, line.rule, line.kept) == (1, "atLeast", False)
    pair = reading(storage_of("w:ind/@w:firstLine"), "firstLine", 32768)
    assert (pair.value, pair.attribute) == (32768, "hanging")
    assert reading(storage_of("w:tab/@w:pos"), "pos", 32767).value == 31680
    assert reading(storage_of("w:pgSz/@w:w"), "w", 65536).value == 144
    assert reading(storage_of("w:col/@w:w"), "w", -5).value is None
    frame = reading(storage_of("w:framePr/@w:h"), "h", 32768, "exact")
    assert (frame.value, frame.rule) == (0, "auto")


# ---------------------------------------------------------------------------
# out of range: the report's fields and semantics
# ---------------------------------------------------------------------------


def _paragraph_with(**ind):
    paragraph = p("x")
    paragraph.p_pr = el.pPr(ind=el.ind(**ind))
    return paragraph


def test_the_out_of_range_finding_is_core_ts_s():
    (finding,) = findings_in(W.Body(content=[_paragraph_with(left=7205001)]))
    assert finding.to_dict() == {
        "code": "value.out_of_range",
        "path": "/w:body/w:p[1]/w:pPr[1]/w:ind[1]/@w:left",
        "element": "w:ind",
        "attribute": "w:left",
        "value": 7205001,
        "word_reads": -3959,
        "storage": "int16",
        "range": [-32768, 32767],
        "message": (
            "/w:body/w:p[1]/w:pPr[1]/w:ind[1]/@w:left=7205001 is past what Word 365 stores there "
            "(int16, -32768 to 32767); Word reads it as -3959"
        ),
    }
    (pair,) = findings_in(_paragraph_with(first_line=32768))
    assert (pair.word_reads, pair.read_as, pair.range) == (32768, "w:hanging", (0, 32767))
    assert pair.message.endswith("Word reads it as 32768 in w:hanging")


def test_a_kept_decimal_is_read_as_word_reads_it_first():
    """CR-021: ``"70000.5"`` is 70000 to Word, past int16; ``"720.5"`` 720, in range; the value is as written."""
    (finding,) = findings_in(_paragraph_with(left="70000.5"))
    assert (finding.value, finding.word_reads) == ("70000.5", 4464)
    assert findings_in(_paragraph_with(left="720.5")) == []
    assert findings_in(_paragraph_with(left="abc")) == []  # not a number Word reads: CR-021's, not this


def _spacing(**values):
    paragraph = p("x")
    paragraph.p_pr = el.pPr(spacing=el.spacing(**values))
    return paragraph


def test_the_read_side_semantics_of_section_4_7():
    """65,535 and -1 in the seven ``"-1"`` rows are in range; the write side's 65,534 cap is no read limit."""
    assert findings_in(_spacing(before=65535)) == []
    assert findings_in(_spacing(after=-1)) == []
    assert findings_in(_spacing(before=65536))[0].word_reads == 0
    assert findings_in(_spacing(before=-2))[0].word_reads == 65534
    section = el.sectPr(pg_mar=el.pgMar(left=65535, right=-1, header=65536))
    assert [(f.attribute, f.value) for f in findings_in(section)] == [("w:header", 65536)]


def test_a_column_s_space_sentinel_and_its_computed_width():
    columns = el.cols(col=[el.col(w=100, space=-1), el.col(w=-5, space=-2)])
    found = findings_in(columns)
    assert [(f.attribute, f.value, f.word_reads) for f in found] == [("w:w", -5, None), ("w:space", -2, -1)]
    assert found[0].message.endswith("Word reads it as a width Word computes")


def test_a_row_height_under_auto_and_an_unmeasured_attribute_are_not_checked():
    row = el.tr(tr_pr=el.trPr(content=[el.trHeight(val=70000, h_rule="auto")]))
    assert findings_in(row) == []
    row = el.tr(tr_pr=el.trPr(content=[el.trHeight(val=70000)]))
    assert [(f.word_reads, f.rule) for f in findings_in(row)] == [(4464, "atLeast")]
    grid = el.tblGrid(grid_col=[el.gridCol(w=7205001)])
    assert integer_limit("w:gridCol/@w:w") is None
    assert findings_in(grid) == []


def test_a_table_width_with_no_row_for_its_type_is_not_checked():
    """CR-022's rows are ``dxa`` and ``pct``: under ``auto``, ``nil`` or no ``w:type`` there is no row (Q3)."""
    for kind in ("auto", "nil", None):
        assert findings_in(el.tblPr(tbl_w=el.tblW(w=70000, type_value=kind))) == []
    assert len(findings_in(el.tblPr(tbl_w=el.tblW(w=70000, type_value="dxa")))) == 1


def test_a_holder_named_by_the_row_is_needed():
    """``w:rPr/w:sz``: a run's size is checked; ``w:sz`` elsewhere is not the row's."""
    run = r("x")
    run.r_pr = el.rPr(sz=el.sz(val=70000))
    assert [f.attribute for f in findings_in(p(run))] == ["w:val"]


def test_the_rule_beside_a_line():
    exact = findings_in(_spacing(line=65535, line_rule="exact"))[0]
    assert (exact.word_reads, exact.rule, exact.storage, exact.range) == (1, "atLeast", "int16, the sign the rule", (1, 32768))
    assert exact.message.endswith("Word reads it as atLeast 1")
    auto = findings_in(_spacing(line=32768))[0]
    assert (auto.word_reads, auto.rule, auto.storage) == (-32768, "auto", "int16")


# ---------------------------------------------------------------------------
# the oracle: Word 365's saves of docx4j CR-066 rounds 5 and 6
# ---------------------------------------------------------------------------

#: Attributes whose element is found under a holder in the case's block (the paragraph's or the run's properties).
_UNDER = {
    "w:spacing/@w:before": ["pPr", "spacing"],
    "w:spacing/@w:after": ["pPr", "spacing"],
    "w:spacing/@w:line": ["pPr", "spacing"],
    "w:spacing/@w:val": ["rPr", "spacing"],
    "w:sz/@w:val": ["rPr", "sz"],
    "w:szCs/@w:val": ["rPr", "szCs"],
    "w:kern/@w:val": ["rPr", "kern"],
    "w:position/@w:val": ["rPr", "position"],
    "w:w/@w:val": ["rPr", "w"],
    "w:bdr/@w:sz": ["rPr", "bdr"],
    "w:bdr/@w:space": ["rPr", "bdr"],
    "w:fitText/@w:val": ["rPr", "fitText"],
}
#: Attributes Word reads in a transitional document and writes back under another name.
_SAVED_AS = {"w:ind/@w:start": "left", "w:ind/@w:end": "right"}
#: What our table leaves unmeasured among the cases (CR-022 section 3.3): core-ts holds these to the objects table
#: instead of to Word's save, and reports some of them; this table reports none.
_UNMEASURED = ("w:gridCol/@w:w", "w:wBefore/@w:w", "w:wAfter/@w:w")


@cache
def _cases() -> list[dict]:
    """The 660 integer cases of rounds 5 and 6: all but the controls (``Z``), CR-021's forms (round 6's ``E``) and
    round 6's numbering (``N``), whose ``w:legacy`` Word 365 drops whole; as docx4j-core-ts counts them."""
    out = []
    for round_name, value_column in (("round5", 3), ("round6", 4)):
        lines = (LIMITS / round_name / "cases.txt").read_text(encoding="utf-8").splitlines()[1:]
        for line in lines:
            columns = line.split("\t")
            case, group, attribute = columns[:3]
            if group in ("Z", "E") or (group == "N" and round_name == "round6"):
                continue
            stem = f"cr066-{group}" if round_name == "round5" else f"cr066-r6-{case if group == 'D' else group}"
            out.append({
                "round": round_name, "id": case, "group": group, "attribute": attribute,
                "value": int(columns[value_column]), "file": LIMITS / round_name / f"{stem}.docx",
                "save": LIMITS / round_name / f"{stem}-word365.docx",
                "part": "word/settings.xml" if group == "D" else "word/document.xml",
            })
    return out


@cache
def _xml(path: Path, part: str) -> etree._Element:
    return etree.fromstring(zipfile.ZipFile(path).read(part))


def _text(node: etree._Element) -> str:
    return "".join(t.text or "" for t in node.iter(f"{{{W_NS}}}t"))


def _block(root: etree._Element, case: str) -> list[etree._Element]:
    """The case's label paragraph and, where the value is there, the table or the section-ending paragraph after it."""
    if etree.QName(root).localname == "settings":
        return [root]
    body = root.find(f"{{{W_NS}}}body")
    children = list(body)
    for index, child in enumerate(children):
        if etree.QName(child).localname == "p" and _text(child).startswith(f"{case} "):
            after = children[index + 1] if index + 1 < len(children) else None
            if after is not None and _holds_value(after):
                return [child, after]
            return [child]
    raise AssertionError(f"no label {case}")


def _holds_value(node: etree._Element) -> bool:
    name = etree.QName(node).localname
    return name == "tbl" or (name == "p" and node.find(f".//{{{W_NS}}}sectPr") is not None and _text(node) == "")


def _element_in(block: list[etree._Element], steps: list[str]) -> etree._Element | None:
    """The first element the steps name in the block, each step the parent of the next (core-ts's ``elementIn``)."""
    for root in block:
        for element in [root, *root.iter(f"{{{W_NS}}}{steps[-1]}")]:
            if etree.QName(element).localname != steps[-1]:
                continue
            node, ok = element.getparent(), True
            for step in reversed(steps[:-1]):
                if node is None or etree.QName(node).localname != step:
                    ok = False
                    break
                node = node.getparent()
            if ok:
                return element
    return None


def _steps(attribute: str) -> tuple[list[str], str, str | None]:
    found = re.fullmatch(r"(.*)/@w:(\w+)(?:\[(\w+)\])?", attribute)
    assert found is not None, attribute
    path, name, rule = found.groups()
    rule = rule if rule in ("auto", "exact", "atLeast") else None  # a bracketed table width type is no rule
    base = path + "/@w:" + name
    steps = _UNDER.get(base) or [step.removeprefix("w:") for step in path.split("/")]
    return steps, name, rule


def _row_path(attribute: str) -> str:
    """A case's attribute as the limits table spells its path: ``w:pBdr/w:top/@w:sz`` is ``w:top/@w:sz``."""
    base = re.sub(r"\[\w+\]$", "", attribute)
    if base == "w:sz/@w:val":
        return "w:rPr/w:sz/@w:val"
    keys = {row.attribute for row in integer_limits()}
    if base in keys:
        return base
    tail = "/".join(base.split("/")[-2:])
    return tail if tail in keys else base


def _storage(attribute: str) -> Storage:
    """The case's storage; ``w:fitText``, which this table leaves unmeasured (Word kept every value), as int32."""
    return STORAGE.get(_row_path(attribute), Storage(32, True))


def _attribute(element: etree._Element | None, name: str) -> int | None:
    if element is None:
        return None
    text = element.get(f"{{{W_NS}}}{name}")
    return None if text is None else int(text)


def _case_of(finding: ValueOutOfRange, case: dict) -> str:
    """The case a finding is in: its path evaluated on the source, climbed to the body's child, then its label."""
    root = _xml(case["file"], case["part"])
    if case["part"] == "word/settings.xml":
        return case["id"]
    element_path, attribute = finding.path.rsplit("/@", 1)
    (element,) = root.xpath(element_path, namespaces=NS)
    assert element.get(f"{{{W_NS}}}{attribute.removeprefix('w:')}") == str(finding.value), finding.path
    node = element
    while etree.QName(node.getparent()).localname != "body":
        node = node.getparent()
    if _holds_value(node):
        node = node.getprevious()
    return _text(node).split(" ")[0]


@cache
def _reported() -> dict[tuple[str, str], ValueOutOfRange]:
    """Each case's finding, from ``misplaced()`` over each probe file, every part read privately."""
    out: dict[tuple[str, str], ValueOutOfRange] = {}
    files = {(case["file"], case["part"]): case for case in _cases()}
    for (path, part_name), case in files.items():
        found = misplaced(load(path))
        assert all(isinstance(f, ValueOutOfRange) for f in found), [f for f in found if not isinstance(f, ValueOutOfRange)]
        assert {f.part_name for f in found} <= {f"/{part_name}"}, "the probe's other parts hold nothing out of range"
        for finding in found:
            key = (str(path), _case_of(finding, case))
            assert key not in out, f"{key}: one finding per case"
            out[key] = finding
    return out


def _word_changed(case: dict) -> tuple[bool, dict] | None:
    """Whether Word's save changed the case's value, its rule or its attribute; None where the save cannot say.

    docx4j-core-ts's reading of Word's save (``test/integer-attributes.test.mjs``), step for step: Word kept the
    value exactly when its save writes it back under the same name and rule, or 65,535 where Word writes it -1; a
    value it left out reads as the default 0; a row height under ``hRule="auto"`` is ignored, not changed.
    """
    steps, name, written_rule = _steps(case["attribute"])
    storage = _storage(case["attribute"])
    rule_attribute = "lineRule" if storage.sign_rule == "lineRule" else "hRule" if storage.sign_rule == "hRule" else None
    if storage.height_flag:
        written_rule = written_rule or "exact"
    saved = _element_in(_block(_xml(case["save"], case["part"]), case["id"]), steps)
    saved_as = _SAVED_AS.get(re.sub(r"\[\w+\]$", "", case["attribute"]), name)
    if storage.pair is not None:
        negative = _attribute(saved, storage.pair[1])
        saved_name, saved_value = (storage.pair[1], negative) if negative is not None else (storage.pair[0], _attribute(saved, storage.pair[0]))
    else:
        saved_name, saved_value = saved_as, _attribute(saved, saved_as)
    saved_rule = (saved.get(f"{{{W_NS}}}{rule_attribute}") or None) if rule_attribute and saved is not None else None
    facts = {"value": saved_value, "name": saved_name, "rule": saved_rule, "rule_attribute": rule_attribute}
    if rule_attribute == "hRule" and written_rule == "auto" and not storage.height_flag:
        assert saved_value is None, f"{case['id']}: a row height under auto, dropped by Word"
        return None
    value = case["value"]
    if saved_value is None:
        return True, facts  # left out: it reads as the default 0
    same = (
        saved_value == value
        and saved_name == saved_as
        and (written_rule is None or saved_rule == written_rule
             or (written_rule == "atLeast" and rule_attribute == "lineRule" and saved_rule == "atLeast"))
    )
    row_rule_same = rule_attribute != "hRule" or storage.height_flag or saved_rule == ("exact" if written_rule == "exact" else None)
    kept = (same and row_rule_same) or (storage.max_written_as is not None and value == 65535)
    return not kept, facts


def test_the_probe_cases_are_core_ts_s_660():
    cases = _cases()
    assert len(cases) == 660
    assert sum(case["round"] == "round5" for case in cases) == 190


def test_the_report_flags_exactly_the_cases_word_s_save_changed():
    """docx4j CR-066 rounds 5 and 6: a finding exactly where Word 365's save changed the value, its rule or its
    attribute, with ``word_reads`` what Word wrote; 413 of core-ts's 432, the other 19 the unmeasured attributes
    core-ts holds to the objects table and this table does not check (section 11)."""
    reported = _reported()
    changed: set[tuple[str, str]] = set()
    compared = left_out = auto = unmeasured = core_ts_unmeasured = 0
    for case in _cases():
        key = (str(case["file"]), case["id"])
        base = re.sub(r"\[\w+\]$", "", case["attribute"])
        if base in _UNMEASURED:
            # no oracle in Word's save; core-ts reports these outside the objects table's range (gridCol the certain
            # range, wBefore and wAfter int16): this table has no row for them, so nothing is reported
            assert key not in reported, key
            unmeasured += 1
            low, high = (-32768, 65535) if base == "w:gridCol/@w:w" else (-32768, 32767)
            core_ts_unmeasured += not low <= case["value"] <= high
            continue
        verdict = _word_changed(case)
        if verdict is None:
            assert key not in reported, f"{key}: a row height under auto is ignored, not out of range"
            auto += 1
            compared += 1  # compared with Word's save, which dropped it, as core-ts counts it
            continue
        was_changed, facts = verdict
        finding = reported.get(key)
        assert (finding is not None) == was_changed, (key, case["attribute"], case["value"], facts, finding)
        if was_changed:
            changed.add(key)
        if facts["value"] is None:
            left_out += 1  # Word left it out: it reads as the default 0
            assert finding.word_reads == 0, (key, finding)
            continue
        compared += 1
        if not was_changed:
            continue
        storage = _storage(case["attribute"])
        if storage.negative == "computed" and finding.word_reads is None:
            assert facts["value"] > 0 and facts["value"] != case["value"], key
            continue
        assert finding.word_reads == facts["value"], (key, finding, facts)
        if storage.pair is not None:
            property_name = _steps(case["attribute"])[1]
            assert finding.read_as == (None if facts["name"] == property_name else f"w:{facts['name']}"), key
        if facts["rule_attribute"] == "lineRule":
            assert finding.rule == facts["rule"], key
        elif facts["rule_attribute"] == "hRule":
            assert ("exact" if finding.rule == "exact" else None) == facts["rule"], key
        assert not finding.range[0] <= case["value"] <= finding.range[1]
    assert set(reported) == changed
    assert (compared, left_out, auto, unmeasured) == (614, 16, 8, 30), "core-ts's 614 compared, 16 left out, 30 stated"
    assert (len(changed), core_ts_unmeasured) == (413, 19)
    assert len(changed) + core_ts_unmeasured == 432, "core-ts's and the objects package's count"


def test_word_s_own_saves_hold_nothing():
    for save in sorted({case["save"] for case in _cases()}):
        assert misplaced(load(save)) == [], save


def test_the_rest_of_the_probes_hold_nothing_out_of_range():
    """Round 6's ``E`` (CR-021's digitless forms, read 0) and ``N`` (``w:legacy``, unmeasured), and the controls."""
    for name in ("cr066-r6-E", "cr066-r6-N"):
        assert misplaced(load(LIMITS / "round6" / f"{name}.docx")) == [], name


def test_cr_065_s_decimals_and_cr_021_s_edges_hold_nothing():
    """docx4j CR-065 round 2's 64 cases (``tests/fixtures/integers/``) and CR-021's edges: every value, read as Word
    reads it lexically, is in range, so nothing is reported, in any part, of the probes or of Word's saves."""
    files = sorted((FIXTURES / "integers").glob("*.docx")) + sorted((FIXTURES / "integers-edges").glob("*.docx"))
    assert len(files) > 30
    cases = (FIXTURES / "integers" / "cases.txt").read_text(encoding="utf-8").splitlines()
    assert len([line for line in cases if line.strip() and not line.startswith("Z")]) == 64
    for path in files:
        assert [f for f in misplaced(load(path)) if isinstance(f, ValueOutOfRange)] == [], path


# ---------------------------------------------------------------------------
# the corpus, and packages
# ---------------------------------------------------------------------------


def _corpus() -> list[Path]:
    suffixes = (".docx", ".docm", ".dotx", ".dotm", ".pptx", ".xlsx")
    return sorted(p for p in (*(ROOT / "samples").rglob("*"), *FIXTURES.rglob("*")) if p.suffix.lower() in suffixes)


def test_the_corpus_holds_nothing_but_the_probes_values():
    """``samples/`` and every fixture: no misplaced content and no cell that does not end with a paragraph; the only
    findings are the out-of-range values of docx4j CR-066's probe files, in their case parts."""
    files = _corpus()
    assert len(files) > 380
    other = []
    for path in files:
        for finding in misplaced(load(path)):
            probe = LIMITS in path.parents and not path.name.endswith("-word365.docx")
            if not (probe and isinstance(finding, ValueOutOfRange)):
                other.append((path.name, finding))
    assert other == []


def test_misplaced_never_marks_a_part_for_re_marshalling():
    package = load(ROOT / "samples" / "Symbols.docx")
    assert misplaced(package) == []
    assert not any(getattr(part, "is_unmarshalled", False) for part in package.parts.parts())


def test_misplaced_takes_a_tree_a_part_or_a_package():
    package = load(LIMITS / "round5" / "cr066-P.docx")
    main = package.main_document_part
    by_package = misplaced(package)
    assert by_package == misplaced(main) == package.misplaced()
    assert {f.part_name for f in by_package} == {"/word/document.xml"}
    by_tree = misplaced(main.contents)
    assert [f.path for f in by_tree] == [f.path for f in by_package]
    assert all(f.part_name is None for f in by_tree)
    assert package.misplaced(unmarshalled_only=True) == by_package  # now unmarshalled
    assert load(LIMITS / "round5" / "cr066-P.docx").misplaced(unmarshalled_only=True) == []


def test_an_address_where_the_content_api_has_one():
    package = create_package()
    body = package.main_document_part.contents.body
    body.content.extend([p("first"), p("second")])
    second = body.content[1]
    second.p_pr = el.pPr(ind=el.ind(left=70000))
    (finding,) = package.misplaced()
    assert finding.address == "body/1"
    assert finding.path == "/w:document/w:body[1]/w:p[2]/w:pPr[1]/w:ind[1]/@w:left"
    second.para_id = "5A2B1C3D"
    assert package.misplaced()[0].address == "w14:5A2B1C3D"
    body.content.append(_table(W.Tc()))
    found = _only(package.misplaced(), CellWithoutEndParagraph)
    assert [f.address for f in found] == ["body/2/0/0"]


# ---------------------------------------------------------------------------
# the save's check (Q5)
# ---------------------------------------------------------------------------


def test_the_check_is_off_by_default_and_on_in_this_suite():
    suite = os.environ.get("DOCX4J_PY_SAVE_CHECK", "1") != "0"
    assert saves_check(None) is suite, "tests/conftest.py turns it on, unless DOCX4J_PY_SAVE_CHECK=0"
    with checking(False):
        assert saves_check(None) is False
        assert saves_check(True) is True
    with checking(True):
        assert saves_check(None) is True
    assert saves_check(False) is False
    # outside any checking() block, as a program outside the suite runs: off
    import contextvars

    from docx4j_py.openpackaging import save

    assert contextvars.Context().run(save.saves_check, None) is False


def _broken_package():
    package = create_package()
    with unchecked():
        package.main_document_part.contents.body.content.append(r("a run in the body"))
    return package


def test_a_save_with_the_check_off_writes_what_is_there():
    package = _broken_package()
    data = package.save(check=False)
    assert b"a run in the body" in zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml")
    assert package.last_save.checked is False and package.last_save.findings == []


def test_misplaced_content_raises_and_nothing_is_written(tmp_path):
    package = _broken_package()
    target = tmp_path / "out.docx"
    target.write_bytes(b"as it was")
    with pytest.raises(SaveCheckError) as raised:
        package.save(target, check=True)
    assert target.read_bytes() == b"as it was", "the check runs before the target is opened"
    error = raised.value
    assert error.code == "target.misplaced"
    assert [f.code for f in error.findings] == ["target.misplaced"]
    assert error.message.startswith(
        "the save's check found /word/document.xml /w:document/w:body[1]/w:r[1]: `Body.content` can't hold an `R`"
    )
    assert error.hint == "a run goes in a paragraph: `p(r(...))`"
    assert error.to_dict()["code"] == "target.misplaced"
    with checking(True), pytest.raises(SaveCheckError):
        package.save()  # check=None: the default of the moment, the suite's


def test_a_value_the_source_held_is_carried_and_one_the_session_set_raises():
    package = load(LIMITS / "round5" / "cr066-P.docx")
    part = package.main_document_part
    held = misplaced(part)
    assert len(held) == 30
    package.save(check=True)  # misplaced() read the part privately: the save re-marshals nothing, so checks nothing
    assert package.last_save.checked is True and package.last_save.findings == []
    part.contents  # noqa: B018 - unmarshalled, adopting the private tree: every value is the source's
    package.save(check=True)
    assert package.last_save.findings == held == package.last_save.out_of_range
    # an edit that moves them leaves them the source's
    body = part.contents.body
    body.content.insert(0, p("a new first paragraph"))
    package.save(check=True)
    assert len(package.last_save.findings) == len(held)
    # a copy is the session's
    from docx4j_py.child import deep_copy

    first = next(f.node for f in held if f.attribute == "w:left")
    copied = deep_copy(first.parent.parent)
    body.content.append(copied)
    with pytest.raises(SaveCheckError) as raised:
        package.save(check=True)
    assert raised.value.code == "value.out_of_range"
    assert [f.attribute for f in raised.value.findings] == ["w:left"]


def test_a_value_set_in_code_raises():
    package = create_package()
    body = package.main_document_part.contents.body
    body.content.append(_paragraph_with(left=70000))
    with pytest.raises(SaveCheckError) as raised:
        package.save(check=True)
    assert raised.value.code == "value.out_of_range"
    assert raised.value.findings[0].word_reads == 4464


def test_a_cell_the_source_held_is_carried_and_one_the_session_broke_raises(tmp_path):
    package = create_package()
    body = package.main_document_part.contents.body
    with unchecked():
        body.content.append(_table(W.Tc(content=[p("a"), tbl([["x"]])])))
    path = tmp_path / "cell.docx"
    package.save(path, check=False)
    loaded = load(path)
    loaded.main_document_part.contents.body.content.append(p("edit elsewhere"))
    loaded.save(check=True)
    assert [f.code for f in loaded.last_save.findings] == ["cell.no_end_paragraph"]
    cells = [f.node for f in loaded.last_save.findings]
    assert len(cells) == 1
    # a cell the session empties is its own
    first_table = loaded.main_document_part.contents.body.content[0]
    first_table.content[0].content.append(W.Tc())
    with pytest.raises(SaveCheckError) as raised:
        loaded.save(check=True)
    assert raised.value.code == "cell.no_end_paragraph"
    assert [f.path for f in raised.value.findings] == ["/w:document/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[2]"]


def test_source_tree_touches_nothing():
    package = load(LIMITS / "round5" / "cr066-P.docx")
    part = package.main_document_part
    fresh = part.source_tree()
    assert fresh is not None and not part.is_unmarshalled and part._private_contents is None
    live = part.contents
    assert part.source_tree() is not live
    assert create_package().main_document_part.source_tree() is None
