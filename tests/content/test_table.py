"""``Table``, ``TableRow`` and ``TableCell``. CR-003 Phase C, section 3.2.

Every mutation is followed by a save, a reload and a read back, which is
section 7's rule, and the two rules of section 4 that land on tables get a test
each: rows and cells descend into row- and cell-level content controls (the
OpenDoPE repeat in ``samples/invoice2013.docx``), and ``header_row_count``
counts **leading** rows only.
"""

from __future__ import annotations

import pytest
from conftest import ROOT, part_bytes, reloaded, sample

from docx4j_py import load
from docx4j_py.model.content import ContentError, Paragraph, Table, TableCell

#: A document with a table inside a cell of another table. Copied from docx4j's
#: layout-fidelity corpus (``table-nested.docx``), Apache-2.0.
NESTED = ROOT / "tests" / "fixtures" / "nested-table.docx"

#: The paragraph that fixture puts in the nested table's second row.
LOREM = (
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod "
    "tempor incididunt ut labore et dolore magna aliqua."
)


@pytest.fixture
def table(new_package):
    """A three-by-three table with values, in a created document."""
    return new_package.body.insert_table(
        3,
        3,
        values=[["Region", "Quarter", "Total"], ["North", "Q1", "120"], ["South", "Q2", "90"]],
    )


# ---------------------------------------------------------------------------
# insert_table
# ---------------------------------------------------------------------------


def test_insert_table_sizes_the_grid_from_the_section(new_package, table):
    assert isinstance(table, Table)
    assert table.row_count == 3
    assert table.rows[0].cell_count == 3
    assert table.values == [
        ["Region", "Quarter", "Total"],
        ["North", "Q1", "120"],
        ["South", "Q2", "90"],
    ]

    # the grid is sized from w:sectPr and sums to the table's own width exactly
    from docx4j_py.model.content.table import writable_width

    assert sum(table.column_widths_twips()) == table.element.tbl_pr.tbl_w.w
    assert abs(sum(table.column_widths_twips()) - writable_width(new_package.body)) <= 2
    assert table.column_widths() == [pytest.approx(w / 20) for w in table.column_widths_twips()]

    back = reloaded(new_package)
    assert back.body.tables[0].values == table.values


def test_insert_table_sets_no_style_unless_one_is_asked_for(new_package):
    plain = new_package.body.insert_table(1, 1)
    assert plain.style_id == ""
    assert plain.style == ""
    assert plain.style_built_in == "Other"

    styled = new_package.body.insert_table(1, 1, style="Table Grid")
    assert styled.style_id == "TableGrid"
    assert styled.style == "Table Grid"
    assert styled.style_built_in == "TableGrid"

    plain.style_built_in = "TableGrid"
    assert plain.style_id == "TableGrid"
    plain.style = ""
    assert plain.style_id == ""


def test_insert_table_at_the_start_and_the_refusals(new_package):
    body = new_package.body
    body.insert_paragraph("after")
    body.insert_table(1, 1, location="Start", values=[["first"]])

    assert body[0].text == "first"
    assert isinstance(body[0], Table)

    with pytest.raises(ContentError) as raised:
        body.insert_table(0, 3)
    assert raised.value.code == "table.empty"
    assert "at least one row" in raised.value.message


def test_values_get_and_set_and_the_cell_value(new_package, table):
    table.values = [["A", "B", "C"], ["D", "E", "F"]]
    assert table.values[0] == ["A", "B", "C"]
    assert table.values[2] == ["South", "Q2", "90"], "the third row was not given values"

    cell = table.cell(2, 0)
    assert isinstance(cell, TableCell)
    assert cell.value == "South"
    cell.value = "East"
    assert cell.text == "East"
    assert len(cell.body) == 1, "setting a value replaces the content with one paragraph"

    assert reloaded(new_package).body.tables[0].values[2] == ["East", "Q2", "90"]


def test_cell_refuses_a_row_or_a_cell_that_is_not_there(table):
    with pytest.raises(ContentError) as raised:
        table.cell(9, 0)
    assert raised.value.code == "table.no_row"
    assert "it has 3" in raised.value.message

    with pytest.raises(ContentError) as raised:
        table.cell(0, 9)
    assert raised.value.code == "table.no_cell"


# ---------------------------------------------------------------------------
# rows
# ---------------------------------------------------------------------------


def test_add_rows_takes_the_tables_columns_and_widths(new_package, table):
    added = table.add_rows(2, values=[["East", "Q3", "30"]])

    assert [row.row_index for row in added] == [3, 4]
    assert table.row_count == 5
    assert table.values[3] == ["East", "Q3", "30"]
    assert table.values[4] == ["", "", ""]
    widths = table.column_widths_twips()
    assert [cell.width for cell in added[0].cells] == [pytest.approx(w / 20) for w in widths]

    at_start = table.add_rows(1, location="Start", values=[["TOP", "", ""]])
    assert at_start[0].row_index == 0
    assert table.values[0] == ["TOP", "", ""]

    assert reloaded(new_package).body.tables[0].row_count == 6

    with pytest.raises(ContentError) as raised:
        table.add_rows(0)
    assert raised.value.code == "table.no_rows"


def test_insert_rows_before_and_after_a_row(new_package, table):
    middle = table.rows[1]
    before = middle.insert_rows(1, location="Before", values=[["B", "", ""]])
    after = middle.insert_rows(1, values=[["A", "", ""]])

    assert [row.values[0] for row in table.rows] == ["Region", "B", "North", "A", "South"]
    assert before[0].row_index == 1
    assert after[0].row_index == 3
    assert reloaded(new_package).body.tables[0].row_count == 5


def test_delete_rows_and_delete_a_row(new_package, table):
    table.delete_rows(1, 2)
    assert table.values == [["Region", "Quarter", "Total"]]

    table.add_rows(1, values=[["gone", "", ""]])
    table.rows[1].delete()
    assert table.row_count == 1
    assert reloaded(new_package).body.tables[0].row_count == 1


def test_header_row_count_counts_leading_rows_only(new_package, table):
    assert table.header_row_count == 0

    table.header_row_count = 2
    assert [row.is_header for row in table.rows] == [True, True, False]
    assert table.header_row_count == 2

    # a header row that is not leading does not count
    table.header_row_count = 0
    table.rows[2].is_header = True
    assert table.header_row_count == 0
    assert table.rows[2].is_header is True

    table.rows[0].is_header = True
    assert table.header_row_count == 1

    back = reloaded(new_package)
    assert back.body.tables[0].header_row_count == 1


def test_table_delete_and_the_change_report(new_package, table):
    assert len(new_package.body) == 1
    table.delete()
    assert len(new_package.body) == 0
    assert new_package.last_change.operation == "delete"
    assert new_package.last_change.addresses == ("body/0",)
    assert reloaded(new_package).body.tables == []


# ---------------------------------------------------------------------------
# cells and bodies
# ---------------------------------------------------------------------------


def test_a_cells_body_carries_the_address_prefix(new_package, table):
    cell = table.cell(1, 1)
    assert cell.address == "body/0/1/1"
    assert cell.body.prefix == "body/0/1/1"
    assert cell.row_index == 1
    assert cell.cell_index == 1
    assert cell.parent_row == table.rows[1]
    assert cell.parent_table == table

    paragraph = cell.paragraphs[0]
    assert paragraph.ordinal == "body/0/1/1/0"
    assert new_package.element_at("body/0/1/1/0") == paragraph

    cell.insert_paragraph("more")
    assert cell.text == "Q1\nmore"
    cell.insert_text(" and more")
    assert cell.text == "Q1\nmore and more"
    assert reloaded(new_package).body.tables[0].cell(1, 1).text == "Q1\nmore and more"


def test_cell_width_is_points(new_package, table):
    cell = table.cell(0, 0)
    assert cell.width == pytest.approx(cell.column_width)
    cell.width = 100
    assert cell.element.tc_pr.tc_w.w == 2000
    assert cell.width == 100.0
    cell.column_width = 50
    assert cell.width == 50.0
    assert reloaded(new_package).body.tables[0].cell(0, 0).width == 50.0


def test_parent_table_cell_is_real_now(new_package, table):
    paragraph = table.cell(2, 2).paragraphs[0]
    assert isinstance(paragraph, Paragraph)
    cell = paragraph.parent_table_cell
    assert isinstance(cell, TableCell)
    assert cell.row_index == 2
    assert cell.cell_index == 2
    assert cell.parent_table.element is table.element

    # a paragraph outside a table has none
    assert new_package.body.insert_paragraph("outside").parent_table_cell is None


def test_a_nested_table_is_the_cells_own():
    package = load(NESTED)
    outer = package.body.tables[0]
    assert outer.address == "body/1"
    assert outer.parent_table_cell is None

    cell = outer.cell(0, 1)
    assert len(cell.tables) == 1
    inner = cell.tables[0]
    assert inner.address == "body/1/0/1/0"
    assert inner.values == [
        ["in a", "in b"],
        [LOREM, "x"],
    ]
    assert inner.parent_table_cell == cell
    # the body's own tables are its own list: the nested one is not in it
    assert len(package.body.tables) == 1


# ---------------------------------------------------------------------------
# section 4: rows descend into row-level content controls
# ---------------------------------------------------------------------------


def test_rows_descend_into_a_row_level_content_control():
    package = sample("invoice2013.docx")
    body = package.body
    repeat = next(c for c in body.content_controls if c.form == "row")
    table = repeat.tables[0]

    # the repeat wraps a w:tr in a w:sdt; the rows are what Word shows
    assert table.row_count == 2
    assert [row.cell_count for row in table.rows] == [4, 4]
    assert table.values[0] == ["productcode", "description", "quantity", "price"]

    # and a new row goes into the table's own content, not into the control
    added = table.add_rows(1, values=[["x", "y", "z", "0"]])
    assert added[0].container is table.element.content
    assert table.row_count == 3
    assert table.values[2] == ["x", "y", "z", "0"]


def test_the_outline_reports_a_repeat_as_one_table():
    package = sample("invoice2013.docx")
    entries = package.outline().entries
    tables = [entry for entry in entries if entry.kind == "table"]
    assert tables, "the invoice has a table"
    shape = next(entry for entry in tables if entry.ordinal == "body/7")
    assert shape.rows == 2, "the row-level w:sdt is unwrapped, not counted as a row"
    assert shape.cols == 4
    assert "productcode" in shape.text


# ---------------------------------------------------------------------------
# the view is a value, and the markdown
# ---------------------------------------------------------------------------


def test_views_compare_equal_hash_and_repr(new_package, table):
    again = new_package.body.tables[0]
    assert again == table
    assert hash(again) == hash(table)
    assert len({again, table}) == 1
    assert repr(table) == "<Table body/0 3x3 'Region | Quarter | Total'>"
    assert repr(table.rows[0]) == "<TableRow 0 of body/0, 3 cells>"
    assert repr(table.cell(0, 0)) == "<TableCell body/0/0/0 'Region'>"
    assert str(table.cell(0, 0)) == "Region"

    assert table.to_dict()["rows"] == 3
    assert table.rows[0].to_dict()["is_header"] is False
    assert table.cell(0, 0).to_dict()["cell_index"] == 0


def test_to_markdown_is_the_renderer_over_the_element(table):
    assert table.to_markdown().splitlines()[0] == "| Region | Quarter | Total |"
    assert table.to_markdown().splitlines()[1] == "| --- | --- | --- |"
    # and it still works over a bare element, which is what Phase K's tests do
    from docx4j_py.model.markdown import table_markdown

    assert table_markdown(table.element) == table.to_markdown()


def test_a_table_in_a_header_body():
    package = sample("Headers.docx")
    header = package.header_parts()[0]
    body = header.body
    added = body.insert_table(1, 2, values=[["left", "right"]])

    assert added.address.startswith("header:rId")
    assert body.tables[0].values == [["left", "right"]]

    data = package.save()
    back = load(data)
    header_back = next(
        part for part in back.header_parts() if part.part_name == header.part_name
    )
    assert header_back.body.tables[0].values == [["left", "right"]]


def test_only_the_touched_part_is_rebuilt():
    package = sample("tables.docx")
    _ = package.body  # the only part the edit touches
    before = part_bytes(package.save())

    package.body.tables[0].add_rows(1, values=[["x", "y"]])
    after = part_bytes(package.save())

    assert after["word/document.xml"] != before["word/document.xml"]
    for name, data in before.items():
        if name != "word/document.xml":
            assert after[name] == data, name
