"""CR-022 D3 and D4: every setter that writes an integer measurement refuses a value Word would misread.

For each measured attribute, through each setter that writes it: both ends of the range the
setter writes within are accepted and one past each is refused, ``MeasureRangeError``
(``value.out_of_range``, a ``ValueError`` too), its hint naming the range, before anything
changes. Which range (D4, Q1): the content API's ``space_before``, ``space_after``,
``line_spacing``, ``List.set_level_indents``' text indent and ``Font.size`` stop where Word
365's own setters stop (core-ts check 71b; CR-024, which ``test_office_js_setters.py`` holds to
Word's run case by case); every other setter, the content
API's indents and widths included, at the faithful range of the attribute's row
(``docx4j_py/resources/integer-limits.json``, which ``tests/content/test_integer_limits.py``
holds to Word's saves); the python-docx facade and the builders at the faithful range
throughout (Q4, Q5). A negative in an attribute Word stores unsigned, or in
``w:trHeight/@w:val``, is refused; one in a signed attribute (the indents, ``w:pgMar`` top and
bottom) is accepted (Q7).

The examples of CR-022 section 2.1 are refused (:func:`test_the_examples_of_section_2_1_are_refused`).
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

import pytest

from docx4j_py import create_package
from docx4j_py.docx import Cm, Document, Pt, Twips
from docx4j_py.model.content import MeasureRangeError
from docx4j_py.model.content.errors import ContentError
from docx4j_py.wml import p, r, tbl, tc, to_xml

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _package() -> Any:
    package = create_package()
    package.id_seed = 20261026
    return package


def refused(write: Callable[[], Any], *in_hint: str) -> MeasureRangeError:
    """`write` raises ``value.out_of_range``, a ``ValueError`` and a ``ContentError``, its hint naming `in_hint`."""
    with pytest.raises(MeasureRangeError) as raised:
        write()
    error = raised.value
    assert error.code == "value.out_of_range"
    assert isinstance(error, ValueError) and isinstance(error, ContentError)
    for text in in_hint:
        assert text in error.hint, (text, error.hint)
        assert text in error.message, (text, error.message)
    return error


# ---------------------------------------------------------------------------
# the content API (points)
# ---------------------------------------------------------------------------

#: (setter, attribute it writes, what is accepted, what is refused, the range as the hint names it)
PARAGRAPH = [
    ("left_indent", "left", (-1638.4, 1638.35, 0, -0.05), (-1638.45, 1638.4, 400000.0), "-1,638.4 to 1,638.35 pt"),
    ("right_indent", "right", (-1638.4, 1638.35), (-1638.45, 1638.4), "-1,638.4 to 1,638.35 pt"),
    ("space_before", "before", (0, 1584), (1584.05, 1638, -0.05, -1, 5000.0), "0 to 1,584 pt"),
    ("space_after", "after", (0, 1584), (1584.05, 3276.75, -0.05), "0 to 1,584 pt"),
    ("line_spacing", "line", (0.75, 1584), (0.7, 0, 1584.05, -1), "0.75 to 1,584 pt"),
]


@pytest.mark.parametrize(("name", "field", "accepted", "rejected", "span"), PARAGRAPH, ids=[row[0] for row in PARAGRAPH])
def test_a_paragraphs_measurement_setters(name: str, field: str, accepted: tuple, rejected: tuple, span: str) -> None:
    package = _package()
    paragraph = package.body.insert_paragraph("x")
    for value in accepted:
        setattr(paragraph, name, value)
        holder = paragraph.element.p_pr.ind if field in ("left", "right") else paragraph.element.p_pr.spacing
        assert getattr(holder, field) == round(value * 20)
    for value in rejected:
        before = to_xml(paragraph.element)
        changes = len(package.changes)
        refused(lambda value=value: setattr(paragraph, name, value), span)
        assert to_xml(paragraph.element) == before, "a refusal changes nothing"
        assert len(package.changes) == changes, "and records no ChangeReport"


def test_the_strict_indent_names_are_held_to_the_same_range() -> None:
    """A paragraph that already uses ``w:start`` and ``w:end`` is written there, within their rows (round 6)."""
    package = _package()
    paragraph = package.body.insert_paragraph("x")
    from docx4j_py.wml import el

    paragraph.element.p_pr = el.pPr(ind=el.ind(start=100, end=100))
    paragraph.left_indent = -1638.4
    paragraph.right_indent = 1638.35
    assert (paragraph.element.p_pr.ind.start, paragraph.element.p_pr.ind.end) == (-32768, 32767)
    assert paragraph.element.p_pr.ind.left is None
    error = refused(lambda: setattr(paragraph, "left_indent", 1638.4), "-1,638.4 to 1,638.35 pt")
    assert error.attribute == "w:ind/@w:start"


def test_the_first_line_indent_is_one_signed_field() -> None:
    """A first line 0 to 1,638.35 pt (32,767 twips), a hanging indent to 1,638.4 (32,768): round 6, check 71."""
    package = _package()
    paragraph = package.body.insert_paragraph("x")
    ind = lambda: paragraph.element.p_pr.ind
    paragraph.first_line_indent = 1638.35
    assert (ind().first_line, ind().hanging) == (32767, None)
    paragraph.first_line_indent = -1638.4
    assert (ind().first_line, ind().hanging) == (None, 32768)
    error = refused(lambda: setattr(paragraph, "first_line_indent", 1638.4), "0 to 1,638.35 pt")
    assert error.attribute == "w:ind/@w:firstLine" and error.stored == 32768
    error = refused(lambda: setattr(paragraph, "first_line_indent", -1638.45), "0 to 1,638.4 pt")
    assert error.attribute == "w:ind/@w:hanging" and error.stored == 32769
    assert (ind().first_line, ind().hanging) == (None, 32768), "the refusals left the last write"


def test_a_line_spacing_is_written_auto_as_words_setter_writes_it() -> None:
    """``line_spacing`` writes 20 x the points under ``lineRule="auto"``, from 0.75 pt (core-ts check 71b; CR-024)."""
    package = _package()
    paragraph = package.body.insert_paragraph("x")
    from docx4j_py.wml import el

    paragraph.element.p_pr = el.pPr(spacing=el.spacing(line=480, line_rule="exact"))
    paragraph.line_spacing = 0.75
    assert (paragraph.element.p_pr.spacing.line, paragraph.element.p_pr.spacing.line_rule) == (15, "auto")
    assert paragraph.line_spacing == 0.75, "the getter reads an auto line as Office JS does"
    paragraph.line_spacing = 24
    assert (paragraph.element.p_pr.spacing.line, paragraph.element.p_pr.spacing.line_rule) == (480, "auto")
    error = refused(lambda: setattr(paragraph, "line_spacing", 0.7), "0.75 to 1,584 pt")
    assert error.stored == 14 and error.attribute == "w:spacing/@w:line"


def test_a_refusal_under_change_tracking_records_nothing() -> None:
    """No ``w:pPrChange``, no report: the check runs before the setter reaches the paragraph's properties."""
    package = _package()
    paragraph = package.body.insert_paragraph("x")
    package.change_tracking_mode = "TrackAll"
    before = to_xml(paragraph.element)
    for name in ("left_indent", "first_line_indent", "space_before", "line_spacing"):
        refused(lambda name=name: setattr(paragraph, name, 400000.0))
    assert to_xml(paragraph.element) == before
    assert "pPrChange" not in to_xml(paragraph.element)


def test_a_cells_width() -> None:
    """``TableCell.width`` and ``column_width``: 0 to 1,638.35 pt; a negative refused, as Office JS refuses it."""
    package = _package()
    table = package.body.insert_table(1, 2)
    cell = table.rows[0].cells[0]
    for setter in ("width", "column_width"):
        for value in (0, 1638.35):
            setattr(cell, setter, value)
            assert cell.element.tc_pr.tc_w.w == round(value * 20)
        before = to_xml(table.element)
        for value in (1638.4, -0.05, 400000.0):
            refused(lambda value=value, setter=setter: setattr(cell, setter, value), "0 to 1,638.35 pt")
        assert to_xml(table.element) == before


def test_a_list_levels_indents() -> None:
    """A text indent short of 1,584 pt either way (Word's own setter, check 71b); a marker within its field (round 6)."""
    package = _package()
    item = package.body.insert_paragraph("one")
    lst = item.start_new_list()
    numbering = package.main_document_part.numbering_definitions_part

    def level_ind() -> Any:
        abstract = next(a for a in numbering.contents.abstract_num if a.abstract_num_id == _abstract_of(lst.id))
        return abstract.lvl[0].p_pr.ind

    def _abstract_of(num_id: int) -> int:
        num = next(n for n in numbering.contents.num if n.num_id == num_id)
        return num.abstract_num_id.val

    lst.set_level_indents(0, -1583.95, -1638.4)
    assert (level_ind().left, level_ind().first_line, level_ind().hanging) == (-31679, None, 32768)
    lst.set_level_indents(0, 1583.95, 1638.35)
    assert (level_ind().left, level_ind().first_line, level_ind().hanging) == (31679, 32767, None)
    changes = len(package.changes)
    before = to_xml(numbering.contents)
    for text in (1584, -1584, 1638, -1638.4):
        refused(lambda text=text: lst.set_level_indents(0, text, 0), "-1,583.95 to 1,583.95 pt")
    refused(lambda: lst.set_level_indents(0, 36, 1638.4), "0 to 1,638.35 pt")
    refused(lambda: lst.set_level_indents(0, 36, -1638.45), "0 to 1,638.4 pt")
    assert to_xml(numbering.contents) == before
    assert len(package.changes) == changes


def test_a_fonts_size() -> None:
    """1 to 1,638 pt after rounding to a half-point, as Word's own setter takes it (check 71b); 0 removes ``w:sz``."""
    package = _package()
    paragraph = package.body.insert_paragraph("x")
    font = paragraph.font
    for value in (0.75, 0.9, 1, 1638, 1638.05):
        font.size = value
        assert paragraph.runs[0].r_pr.sz.val == round(value * 2)
    before = to_xml(paragraph.element)
    changes = len(package.changes)
    for value in (0.5, 0.7, 1638.4, 1638.5, 32767.5, -1, -0.5):
        refused(lambda value=value: setattr(font, "size", value), "1 to 1,638 pt")
    assert to_xml(paragraph.element) == before
    assert len(package.changes) == changes
    font.size = 0
    assert paragraph.runs[0].r_pr.sz is None and paragraph.runs[0].r_pr.sz_cs is None
    font.size = None  # leaves the size alone, as before
    assert paragraph.runs[0].r_pr.sz is None


def test_a_ranges_font_splits_nothing_when_refused() -> None:
    """A range's font splits its runs at a write; a refused size splits none."""
    package = _package()
    paragraph = package.body.insert_paragraph("abcdef")
    span = paragraph.search("cd")[0]
    before = to_xml(paragraph.element)
    refused(lambda: setattr(span.font, "size", 2000))
    assert to_xml(paragraph.element) == before
    span.font.size = 12
    assert len(paragraph.runs) == 3


def test_a_value_that_is_not_a_finite_number_is_refused() -> None:
    package = _package()
    paragraph = package.body.insert_paragraph("x")
    for value in (float("nan"), float("inf"), float("-inf")):
        error = refused(lambda value=value: setattr(paragraph, "left_indent", value), "-1,638.4 to 1,638.35 pt")
        assert "is not a finite number" in error.message


# ---------------------------------------------------------------------------
# the python-docx facade (Length)
# ---------------------------------------------------------------------------

#: (the facade's setter on a section, the attribute, (low, high) in twips, the hint's range in EMU). The five
#: unsigned margins stop at 65,534: Word keeps 0xFFFF but writes it back as -1 (CR-022 section 10, decision 3).
SECTION = [
    ("page_width", "w", (144, 31681), "91,440 to 20,117,435 EMU"),
    ("page_height", "h", (144, 65535), "91,440 to 41,614,725 EMU"),
    ("top_margin", "top", (-32768, 32767), "-20,807,680 to 20,807,045 EMU"),
    ("bottom_margin", "bottom", (-32768, 32767), "-20,807,680 to 20,807,045 EMU"),
    ("left_margin", "left", (0, 65534), "0 to 41,614,090 EMU"),
    ("right_margin", "right", (0, 65534), "0 to 41,614,090 EMU"),
    ("header_distance", "header", (0, 65534), "0 to 41,614,090 EMU"),
    ("footer_distance", "footer", (0, 65534), "0 to 41,614,090 EMU"),
    ("gutter", "gutter", (0, 65534), "0 to 41,614,090 EMU"),
]


@pytest.mark.parametrize(("name", "field", "bounds", "span"), SECTION, ids=[row[0] for row in SECTION])
def test_a_sections_measurements(name: str, field: str, bounds: tuple[int, int], span: str) -> None:
    document = Document()
    section = document.sections[0]
    holder = (lambda: section._sect_pr.pg_sz) if name.startswith("page_") else (lambda: section._sect_pr.pg_mar)
    low, high = bounds
    for twips in (low, high):
        setattr(section, name, Twips(twips))
        assert getattr(holder(), field) == twips
    before = to_xml(section._sect_pr)
    changes = len(document.package.changes)
    for twips in (low - 1, high + 1):
        refused(lambda twips=twips: setattr(section, name, Twips(twips)), span)
    assert to_xml(section._sect_pr) == before
    assert len(document.package.changes) == changes
    setattr(section, name, None)  # None removes the attribute, as before
    assert getattr(holder(), field) is None


@pytest.mark.parametrize(
    ("name", "field", "bounds"),
    [
        ("left_indent", "left", (-32768, 32767)),
        ("right_indent", "right", (-32768, 32767)),
        ("space_before", "before", (0, 65534)),
        ("space_after", "after", (0, 65534)),
    ],
)
def test_a_paragraph_formats_measurements(name: str, field: str, bounds: tuple[int, int]) -> None:
    document = Document()
    paragraph = document.add_paragraph("x")
    form = paragraph.paragraph_format
    holder = (lambda: paragraph._p.p_pr.ind) if "indent" in name else (lambda: paragraph._p.p_pr.spacing)
    low, high = bounds
    for twips in (low, high):
        setattr(form, name, Twips(twips))
        assert getattr(holder(), field) == twips
    before = to_xml(paragraph._p)
    for twips in (low - 1, high + 1):
        refused(lambda twips=twips: setattr(form, name, Twips(twips)))
    assert to_xml(paragraph._p) == before
    setattr(form, name, None)
    assert getattr(holder(), field, None) is None


def test_65535_twips_is_refused_where_word_writes_it_back_as_minus_1() -> None:
    """Space before and after and the five unsigned margins stop at 65,534 twips (decision 3); the page height doesn't."""
    document = Document()
    form = document.add_paragraph("x").paragraph_format
    section = document.sections[0]
    for owner, name in ((form, "space_before"), (form, "space_after"), (section, "left_margin"), (section, "gutter")):
        error = refused(lambda owner=owner, name=name: setattr(owner, name, Twips(65535)), "0 to 41,614,090 EMU")
        assert error.stored == 65535
        setattr(owner, name, Twips(65534))
    section.page_height = Twips(65535)  # Word writes 65535 back as written here
    assert section._sect_pr.pg_sz.h == 65535


def test_a_paragraph_formats_first_line_indent() -> None:
    document = Document()
    paragraph = document.add_paragraph("x")
    form = paragraph.paragraph_format
    form.first_line_indent = Twips(32767)
    assert (paragraph._p.p_pr.ind.first_line, paragraph._p.p_pr.ind.hanging) == (32767, None)
    form.first_line_indent = Twips(-32768)
    assert (paragraph._p.p_pr.ind.first_line, paragraph._p.p_pr.ind.hanging) == (None, 32768)
    before = to_xml(paragraph._p)
    refused(lambda: setattr(form, "first_line_indent", Twips(32768)))
    refused(lambda: setattr(form, "first_line_indent", Twips(-32769)))
    assert to_xml(paragraph._p) == before


def test_a_paragraph_formats_line_spacing_by_its_rule() -> None:
    """A Length exact (1 to 32,768 twips) or at least (0 to 32,767); a multiple in 240ths (-32,768 to 32,767)."""
    document = Document()
    paragraph = document.add_paragraph("x")
    form = paragraph.paragraph_format
    spacing = lambda: paragraph._p.p_pr.spacing
    form.line_spacing = Twips(32768)
    assert (spacing().line, spacing().line_rule) == (32768, "exact")
    form.line_spacing = Twips(1)
    refused(lambda: setattr(form, "line_spacing", Twips(32769)))
    refused(lambda: setattr(form, "line_spacing", Twips(0)))
    spacing().line_rule = "atLeast"
    form.line_spacing = Twips(0)
    form.line_spacing = Twips(32767)
    assert (spacing().line, spacing().line_rule) == (32767, "atLeast")
    refused(lambda: setattr(form, "line_spacing", Twips(32768)))
    form.line_spacing = 32767 / 240
    assert (spacing().line, spacing().line_rule) == (32767, "auto")
    form.line_spacing = -32768 / 240
    assert spacing().line == -32768
    error = refused(lambda: setattr(form, "line_spacing", 137.0), "-136.533 to 136.529 lines")
    assert error.stored == 32880


def test_a_line_spacing_rule_rechecks_the_line_already_set() -> None:
    """A rule over a line already set holds that line to the new rule's range (CR-022 section 10, decision 4).

    Exact is 1 to 32,768 twips, at least 0 to 32,767, auto (or no rule) -32,768 to 32,767 240ths;
    a multiple writes its own line. Refused before anything changes, the hint naming the value and
    the rule.
    """
    from docx4j_py.docx import WD_LINE_SPACING

    document = Document()
    paragraph = document.add_paragraph("x")
    form = paragraph.paragraph_format
    spacing = lambda: paragraph._p.p_pr.spacing
    form.line_spacing = -100 / 240  # an auto line of -100 240ths
    assert (spacing().line, spacing().line_rule) == (-100, "auto")
    before = to_xml(paragraph._p)
    changes = len(document.package.changes)
    error = refused(lambda: setattr(form, "line_spacing_rule", WD_LINE_SPACING.EXACTLY), "1 to 32,768 twips")
    assert error.message == (
        'line spacing, under w:lineRule="exact", -100 twips is beyond Word\'s range, 1 to 32,768 twips'
    )
    assert error.hint == (
        'the line spacing already set, -100 twips, is beyond the range of w:lineRule="exact", 1 to 32,768 twips: '
        "set line_spacing within it first, or choose another rule"
    )
    refused(lambda: setattr(form, "line_spacing_rule", WD_LINE_SPACING.AT_LEAST), "0 to 32,767 twips")
    assert to_xml(paragraph._p) == before, "nothing changed"
    assert len(document.package.changes) == changes, "and no report"
    form.line_spacing_rule = WD_LINE_SPACING.MULTIPLE  # auto keeps -100, which it admits
    form.line_spacing_rule = WD_LINE_SPACING.DOUBLE  # a multiple writes its own line: nothing to check
    assert (spacing().line, spacing().line_rule) == (480, "auto")
    form.line_spacing = Twips(32768)  # exact, its range's top
    error = refused(lambda: setattr(form, "line_spacing_rule", WD_LINE_SPACING.AT_LEAST), 'w:lineRule="atLeast"')
    assert error.stored == 32768
    error = refused(lambda: setattr(form, "line_spacing_rule", None), "no w:lineRule (auto)")
    assert "-32,768 to 32,767 240ths of a line" in error.hint
    assert (spacing().line, spacing().line_rule) == (32768, "exact")
    form.line_spacing = Twips(32767)
    form.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    form.line_spacing_rule = None
    assert (spacing().line, spacing().line_rule) == (32767, None)


def test_a_rows_height_rule_rechecks_the_height_already_set() -> None:
    """At least (or no rule) 0 to 32,767 twips, exact 1 to 32,768; auto not checked, as Word discards it."""
    from docx4j_py.docx import WD_ROW_HEIGHT
    from docx4j_py.docx.table import _tr_pr_child

    document = Document()
    row = document.add_table(1, 1).rows[0]
    height = lambda: _tr_pr_child(row._tr, "trHeight")
    row.height = Twips(0)
    before = to_xml(row._tr)
    error = refused(lambda: setattr(row, "height_rule", WD_ROW_HEIGHT.EXACTLY), "1 to 32,768 twips")
    assert error.hint == (
        'the row height already set, 0 twips, is beyond the range of w:hRule="exact", 1 to 32,768 twips: '
        "set height within it first, or choose another rule"
    )
    assert to_xml(row._tr) == before
    row.height = Twips(32767)
    row.height_rule = WD_ROW_HEIGHT.EXACTLY
    row.height = Twips(32768)
    refused(lambda: setattr(row, "height_rule", WD_ROW_HEIGHT.AT_LEAST), 'w:hRule="atLeast"')
    refused(lambda: setattr(row, "height_rule", None), "no w:hRule (atLeast)")
    assert (height().val, height().h_rule) == (32768, "exact")
    row.height_rule = WD_ROW_HEIGHT.AUTO  # Word discards a height under auto at every value
    assert height().h_rule == "auto"
    table = document.add_table(1, 1)
    table.rows[0].height_rule = WD_ROW_HEIGHT.EXACTLY  # no height yet: nothing to check
    assert _tr_pr_child(table.rows[0]._tr, "trHeight").val is None


def test_a_tab_stops_position() -> None:
    """``w:tab/@w:pos``: -31,680 to 31,680 twips, where Word clamps (22 in)."""
    document = Document()
    paragraph = document.add_paragraph("x")
    stops = paragraph.paragraph_format.tab_stops
    stop = stops.add_tab_stop(Twips(-31680))
    stops.add_tab_stop(Twips(31680))
    assert [s.position.twips for s in stops] == [-31680, 31680]
    before = to_xml(paragraph._p)
    refused(lambda: stops.add_tab_stop(Twips(31681)), "-20,116,800 to 20,116,800 EMU")
    refused(lambda: setattr(stop, "position", Twips(-31681)))
    assert to_xml(paragraph._p) == before


def test_a_fonts_size_through_the_facade() -> None:
    """``w:sz``: 0 to 65,535 half-points (32,767.5 pt); Pt(0) writes 0 as python-docx does; None removes it."""
    document = Document()
    run = document.add_paragraph().add_run("x")
    for points, half in ((32767.5, 65535), (0, 0)):
        run.font.size = Pt(points)
        assert (run._r.r_pr.sz.val, run._r.r_pr.sz_cs.val) == (half, half)
    before = to_xml(run._r)
    refused(lambda: setattr(run.font, "size", Pt(32768)), "0 to 416,147,250 EMU (0 to 32,767.5 pt)")
    refused(lambda: setattr(run.font, "size", Pt(-0.5)))
    assert to_xml(run._r) == before
    run.font.size = None
    assert run._r.r_pr.sz is None


def test_a_rows_height_by_its_rule() -> None:
    """At least (or no rule) 0 to 32,767 twips; exact 1 to 32,768; a negative refused (Q7); unchecked under auto."""
    document = Document()
    table = document.add_table(1, 1)
    row = table.rows[0]
    row.height = Twips(0)
    row.height = Twips(32767)
    refused(lambda: setattr(row, "height", Twips(32768)))
    refused(lambda: setattr(row, "height", Twips(-1)))
    from docx4j_py.docx import WD_ROW_HEIGHT

    row.height_rule = WD_ROW_HEIGHT.EXACTLY  # over 32,767, which exact admits too
    row.height = Twips(32768)
    row.height = Twips(1)
    refused(lambda: setattr(row, "height", Twips(0)))
    refused(lambda: setattr(row, "height", Twips(32769)))
    row.height_rule = WD_ROW_HEIGHT.AUTO
    row.height = Twips(70000)  # Word discards a height under auto at every value: no row, not checked
    assert row.height.twips == 70000
    from docx4j_py.docx.table import _tr_pr_child

    _tr_pr_child(row._tr, "trHeight").h_rule = "Exact"  # outside the enumeration, kept as text: unset, at least
    refused(lambda: setattr(row, "height", Twips(32768)))
    row.height = Twips(32767)


def test_a_cells_and_a_columns_width_through_the_facade() -> None:
    """``w:tcW`` in dxa: -32,768 to 32,767 twips, signed in Word; ``w:gridCol`` alone is not checked."""
    document = Document()
    table = document.add_table(1, 1)
    cell = table.cell(0, 0)
    for twips in (-32768, 32767):
        cell.width = Twips(twips)
        assert cell._tc.tc_pr.tc_w.w == twips
    before = to_xml(table._tbl)
    refused(lambda: setattr(cell, "width", Twips(32768)))
    refused(lambda: setattr(cell, "width", Twips(-32769)))
    refused(lambda: table.add_column(Twips(32768)))
    assert to_xml(table._tbl) == before
    table.add_column(Twips(32767))
    table.columns[0].width = Twips(70000)  # w:gridCol: Word recomputes the grid; not checked (Q3)


def test_add_table_holds_each_columns_width() -> None:
    """``add_table(rows, cols, width)`` writes ``width / cols`` on every cell: each within ``w:tcW``'s range.

    Checked before the table goes in, and in ``_Cell.add_table`` before the paragraph it puts in
    first: a cell whose own width is already beyond the range (a file's value, CR-020's to
    report) refuses a table in it.
    """
    document = Document()
    header = document.sections[0].header
    made = header.add_table(1, 2, Twips(65534))
    assert [cell.tc_pr.tc_w.w for cell in made._tbl.content[0].content] == [32767, 32767]
    tables = len(header.tables)
    refused(lambda: header.add_table(1, 2, Twips(65536)), "-32,768 to 32,767 twips")
    assert len(header.tables) == tables, "nothing was added"
    cell = made.cell(0, 0)
    content = len(cell._tc.content)
    cell._tc.tc_pr.tc_w.w = 70000
    refused(lambda: cell.add_table(1, 2))
    assert len(cell._tc.content) == content, "the paragraph cell.add_table puts in first is not put in"


# ---------------------------------------------------------------------------
# the builders (points for a size, twips for a width)
# ---------------------------------------------------------------------------


def test_the_builders_size() -> None:
    """``r(..., size=)`` and ``p(..., size=)``: 0 to 32,767.5 pt, as Word stores ``w:sz``; 0 writes no size."""
    assert r("x", size=32767.5).r_pr.sz.val == 65535
    assert r("x", size=0.5).r_pr.sz.val == 1
    assert r("x", size=0).r_pr is None or r("x", size=0).r_pr.sz is None
    refused(lambda: r("x", size=32768), "0 to 32,767.5 pt")
    refused(lambda: r("x", size=-0.5), "0 to 32,767.5 pt")
    refused(lambda: p("x", size=40000), "0 to 32,767.5 pt")
    refused(lambda: r("x", bold=True, size=40000))


def test_the_builders_widths() -> None:
    """``tc(width=)``, ``tbl(widths=)`` and ``tbl(width=)``: each cell and the table within -32,768 to 32,767 twips."""
    assert tc("x", width=32767).tc_pr.tc_w.w == 32767
    assert tc("x", width=-32768).tc_pr.tc_w.w == -32768
    refused(lambda: tc("x", width=32768), "-32,768 to 32,767 twips")
    refused(lambda: tc("x", width=-32769), "-32,768 to 32,767 twips")
    table = tbl([["a", "b"]], widths=[16000, 16767])
    assert table.tbl_pr.tbl_w.w == 32767
    refused(lambda: tbl([["a", "b"]], widths=[16000, 16768]), "-32,768 to 32,767 twips")
    refused(lambda: tbl([["a"]], widths=[40000]))
    refused(lambda: tbl([["a", "b"]], width=65536))
    assert tbl([["a", "b"]], width=32767).tbl_pr.tbl_w.w == 32767


# ---------------------------------------------------------------------------
# CR-022 section 2.1
# ---------------------------------------------------------------------------


def test_the_examples_of_section_2_1_are_refused() -> None:
    """What saved ``w:left="8000000"``, ``w:before="100000"``, ``w:pgSz w:w="283465"`` and ``w:left="566929"``."""
    package = _package()
    paragraph = package.body.insert_paragraph("x")
    error = refused(lambda: setattr(paragraph, "left_indent", 400000.0))
    assert error.message == "left indent 400000 pt (8,000,000 twips) is beyond Word's range, -1,638.4 to 1,638.35 pt"
    assert error.hint == "pass a left indent from -1,638.4 to 1,638.35 pt"
    assert (error.attribute, error.value, error.stored) == ("w:ind/@w:left", 400000.0, 8000000)
    error = refused(lambda: setattr(paragraph, "space_before", 5000.0))
    assert error.message == (
        "space before 5000 pt (100,000 twips) is beyond Word's range, 0 to 1,584 pt, "
        "where Office JS's setter stops"
    )
    assert paragraph.element.p_pr is None, "nothing was written"

    document = Document()
    error = refused(lambda: setattr(document.sections[0], "page_width", Cm(500)))
    assert error.message == (
        "page width 180000000 EMU (500 cm, 283,465 twips) is beyond Word's range, "
        "91,440 to 20,117,435 EMU (0.254 to 55.881 cm)"
    )
    paragraph_format = document.add_paragraph("x").paragraph_format
    error = refused(lambda: setattr(paragraph_format, "left_indent", Cm(1000)))
    assert error.stored == 566929
    assert error.hint == "pass a left indent from -20,807,680 to 20,807,045 EMU (-57.799 to 57.797 cm)"
    assert error.to_dict()["code"] == "value.out_of_range"
