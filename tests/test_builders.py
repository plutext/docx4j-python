"""The text-in sugar, CR-001 section 6.2: ``p``, ``r``, ``t``, ``tbl``, ``br``, ``tab``.

Mirrors ``docx4j-generated-objects-ts``'s ``test/smoke.mjs`` for
``builders/wml``: the run-option mapping round-trips, the removal values remove,
``t`` sets ``xml:space``, ``r`` turns a tab and a newline into ``w:tab`` and
``w:br``, and ``tbl`` builds its grid.

    .venv-fork/bin/python -m pytest tests/test_builders.py -q
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py import wml as W
from docx4j_py.wml import el, to_xml
from docx4j_py.wml.builders import (
    UNDERLINE,
    apply_run_options,
    br,
    highlight_hex_value,
    highlight_name_for_color,
    p,
    r,
    read_run_options,
    t,
    tab,
    tbl,
)

# ---------------------------------------------------------------------------
# t, br, tab
# ---------------------------------------------------------------------------


def test_t_sets_xml_space_like_docx4j():
    assert t("Hello").space is None
    assert t(" leading").space == "preserve"
    assert t("trailing ").space == "preserve"
    assert t("two  spaces").space == "preserve"
    assert to_xml(t("  x  ")).endswith('xml:space="preserve">  x  </w:t>')


def test_br_and_tab():
    assert br().type_value is None
    assert br("textWrapping").type_value is None
    assert br("page").type_value == "page"
    assert to_xml(br("page")).startswith("<w:br ")
    assert isinstance(tab(), W.RTab)
    assert to_xml(tab()).startswith("<w:tab ")


# ---------------------------------------------------------------------------
# r
# ---------------------------------------------------------------------------


def test_r_of_text():
    run = r("hello")
    assert isinstance(run, W.R)
    assert run.r_pr is None
    assert [type(item).__name__ for item in run.content] == ["RT"]


def test_r_splits_tabs_and_newlines():
    run = r("a\tb\nc")
    assert [type(item).__name__ for item in run.content] == [
        "RT",
        "RTab",
        "RT",
        "Br",
        "RT",
    ]
    assert to_xml(run).count("<w:tab/>") == 1
    assert to_xml(run).count("<w:br/>") == 1


def test_r_takes_objects_as_they_are():
    drawing = el.drawing()
    run = r("before", drawing, "after")
    assert run.content[1] is drawing
    assert drawing.parent is run


def test_r_run_options_land_on_rpr():
    run = r("x", bold=True, italic=True, size=12, color="#FF0000", font="Calibri")
    assert run.r_pr.b is not None and run.r_pr.b_cs is not None
    assert run.r_pr.i is not None and run.r_pr.i_cs is not None
    assert run.r_pr.sz.val == 24 and run.r_pr.sz_cs.val == 24
    assert run.r_pr.color.val == "FF0000"
    assert run.r_pr.r_fonts.ascii == "Calibri" and run.r_pr.r_fonts.h_ansi == "Calibri"


def test_r_rejects_an_unknown_option():
    with pytest.raises(TypeError, match="unknown run option"):
        r("x", weight="bold")


# ---------------------------------------------------------------------------
# the run-option mapping
# ---------------------------------------------------------------------------

ALL_OPTIONS = {
    "bold": True,
    "italic": True,
    "underline": "Double",
    "strike_through": True,
    "double_strike_through": True,
    "superscript": True,
    "name": "Georgia",
    "size": 11.5,
    "color": "#123ABC",
    "highlight_color": "#FFFF00",
    "style": "Strong",
}


def test_apply_then_read_round_trips_every_option():
    read = read_run_options(apply_run_options(W.RPr(), **ALL_OPTIONS))
    assert read["bold"] is True
    assert read["italic"] is True
    assert read["underline"] == "Double"
    assert read["strike_through"] is True
    assert read["double_strike_through"] is True
    assert read["superscript"] is True
    assert read["subscript"] is False
    assert read["name"] == "Georgia"
    assert read["size"] == 11.5
    assert read["color"] == "#123ABC"
    assert read["highlight_color"] == "#FFFF00"
    assert read["style"] == "Strong"


def test_underline_true_is_single_and_false_is_none():
    assert (
        read_run_options(apply_run_options(W.RPr(), underline=True))["underline"]
        == "Single"
    )
    assert (
        read_run_options(apply_run_options(W.RPr(), underline=False))["underline"]
        == "None"
    )
    assert apply_run_options(W.RPr(), underline=False).u is None


def test_every_underline_name_maps_and_comes_back():
    for name in UNDERLINE:
        read = read_run_options(apply_run_options(W.RPr(), underline=name))
        assert read["underline"] == name


def test_removal_values_remove():
    r_pr = apply_run_options(W.RPr(), **ALL_OPTIONS)
    apply_run_options(
        r_pr,
        bold=False,
        italic=False,
        strike_through=False,
        double_strike_through=False,
        superscript=False,
        name="",
        color="",
        highlight_color=None,
        style="",
        underline="None",
    )
    assert (r_pr.b, r_pr.i, r_pr.strike, r_pr.dstrike, r_pr.u) == (
        None,
        None,
        None,
        None,
        None,
    )
    assert (r_pr.r_fonts, r_pr.color, r_pr.highlight, r_pr.r_style) == (
        None,
        None,
        None,
        None,
    )
    assert r_pr.vert_align is None


def test_color_auto_and_case():
    assert apply_run_options(W.RPr(), color="auto").color.val == "auto"
    assert apply_run_options(W.RPr(), color="#abcdef").color.val == "ABCDEF"


def test_highlight_colour_table():
    assert highlight_name_for_color("#FFFF00") == "yellow"
    assert highlight_name_for_color("yellow") == "yellow"
    assert highlight_name_for_color("#FFD700") == "darkYellow"
    assert highlight_name_for_color("rgb(255,255,0)") is None
    assert highlight_hex_value("cyan") == "#00FFFF"
    assert highlight_hex_value("nope") is None
    with pytest.raises(ValueError, match="not a highlight colour"):
        apply_run_options(W.RPr(), highlight_color="#010203")


def test_read_run_options_of_nothing_is_all_defaults():
    read = read_run_options(None)
    assert read == {
        "bold": False,
        "italic": False,
        "underline": "None",
        "strike_through": False,
        "double_strike_through": False,
        "subscript": False,
        "superscript": False,
        "name": "",
        "size": 0,
        "color": "",
        "highlight_color": None,
        "style": "",
    }


def test_r_produces_what_apply_run_options_does():
    assert r("x", **ALL_OPTIONS).r_pr == apply_run_options(W.RPr(), **ALL_OPTIONS)


def test_font_is_an_alias_of_name():
    assert r("x", font="Arial").r_pr == r("x", name="Arial").r_pr


# ---------------------------------------------------------------------------
# p
# ---------------------------------------------------------------------------


def test_p_of_text_makes_the_run_and_the_text():
    para = p("Hello")
    assert isinstance(para, W.P)
    assert para.p_pr is None
    assert to_xml(para).endswith("<w:r><w:t>Hello</w:t></w:r></w:p>")


def test_p_style_is_the_paragraph_style():
    para = p("Hello", style="Heading1")
    assert para.p_pr.p_style.val == "Heading1"
    assert para.content[0].r_pr is None


def test_p_run_options_apply_to_the_text_and_run_style_is_separate():
    para = p("Hello", style="Heading1", bold=True, run_style="Strong")
    assert para.p_pr.p_style.val == "Heading1"
    assert para.content[0].r_pr.b is not None
    assert para.content[0].r_pr.r_style.val == "Strong"


def test_p_of_runs():
    a, b = r("one ", bold=True), r("two")
    para = p(a, b)
    assert list(para.content) == [a, b]
    assert a.parent is para


def test_p_takes_any_ppr_field():
    para = p("x", keep_next=el.keepNext(), jc=el.jc(val="center"))
    assert para.p_pr.keep_next is not None
    assert para.p_pr.jc.val == "center"


# ---------------------------------------------------------------------------
# tbl
# ---------------------------------------------------------------------------


def test_tbl_grid_from_equal_columns():
    table = tbl([["a", "b", "c"]])
    widths = [c.w for c in table.tbl_grid.grid_col]
    assert widths == [3008, 3008, 3010]
    assert sum(widths) == 9026 == table.tbl_pr.tbl_w.w


def test_tbl_grid_from_given_widths_and_a_style():
    table = tbl([["a", "b"]], style="TableGrid", widths=[1000, 2000])
    assert [c.w for c in table.tbl_grid.grid_col] == [1000, 2000]
    assert table.tbl_pr.tbl_w.w == 3000
    assert table.tbl_pr.tbl_style.val == "TableGrid"


def test_tbl_cells_are_paragraphs_and_short_rows_are_padded():
    table = tbl([["a", "b"], ["c"]])
    assert len(table.content) == 2
    for row in table.content:
        assert isinstance(row, W.Tr)
        assert len(row.content) == 2
        for cell in row.content:
            assert isinstance(cell, W.Tc)
            assert isinstance(cell.content[0], W.P)
    # the padded cell is an empty paragraph with an empty run, as `tbl` in
    # builders/wml.mts makes it
    padded = table.content[1].content[1].content[0]
    assert [type(item).__name__ for item in padded.content] == ["R"]
    assert list(padded.content[0].content) == []


def test_tbl_takes_objects_for_cells():
    para = p("given")
    table = tbl([[para, [p("one"), p("two")]]])
    assert table.content[0].content[0].content[0] is para
    assert len(table.content[0].content[1].content) == 2


def test_tbl_parents_are_linked():
    table = tbl([["a"]])
    row = table.content[0]
    assert row.parent is table
    assert row.content[0].parent is row
    assert row.content[0].content[0].parent is row.content[0]
