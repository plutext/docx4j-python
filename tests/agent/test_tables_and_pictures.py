"""Tool-shaped sessions over the Phase C views. CR-003 section 7.

The same shape as the other files here: every step is something an MCP server
would expose as one call, with JSON-ready arguments and a JSON-ready result.
What is new in Phase C is that a *cell* has an address, that an insert can add
a part, and that ``outline()`` now builds a table's shape from the ``Table``
view.
"""

from __future__ import annotations

import zipfile

import pytest
from conftest import SAMPLES, SEED, big_document, part_bytes, reloaded, sample

from docx4j_py.model.content import Table


@pytest.fixture(scope="module")
def png() -> bytes:
    with zipfile.ZipFile(SAMPLES / "Images.docx") as archive:
        name = next(n for n in archive.namelist() if n.lower().endswith(".png"))
        return archive.read(name)


# ---------------------------------------------------------------------------
# addressing into a table
# ---------------------------------------------------------------------------


def test_a_paragraph_in_a_cell_is_addressed_edited_and_reported(new_package):
    body = new_package.body
    body.insert_paragraph("Quarterly report", style="Heading 1")
    table = body.insert_table(
        2, 2, values=[["Region", "Total"], ["North", "120"]], style="TableGrid"
    )
    new_package.changes.clear()

    # a server holds the address, not the object
    address = table.cell(1, 1).paragraphs[0].address
    assert address == "body/1/1/1/0"

    paragraph = new_package.paragraph_at(address)
    assert paragraph.text == "120"
    paragraph.insert_text(" (up 12%)")

    change = new_package.last_change
    assert change.operation == "insert_text"
    assert change.addresses == (address,)
    assert change.text_before == "120"
    assert change.text_after == "120 (up 12%)"
    assert change.parts_touched == ("/word/document.xml",)
    assert change.moved == ()

    back = reloaded(new_package)
    assert back.paragraph_at(address).text == "120 (up 12%)"
    assert back.body.tables[0].values[1] == ["North", "120 (up 12%)"]


def test_a_table_reports_the_address_it_was_inserted_at(new_package):
    body = new_package.body
    body.insert_paragraph("one")
    table = body.insert_table(1, 1)

    assert isinstance(table, Table)
    assert new_package.last_change.operation == "insert_table"
    assert new_package.last_change.addresses == ("body/1",)
    assert table.address == "body/1"

    # and an insert at the start says which ordinals moved
    new_package.changes.clear()
    body.insert_table(1, 1, location="Start")
    change = new_package.last_change
    assert change.addresses == ("body/0",)
    assert change.moved == (("body/0", "body/1"), ("body/1", "body/2"))


# ---------------------------------------------------------------------------
# dry runs
# ---------------------------------------------------------------------------


def test_a_dry_run_of_insert_table_previews_without_committing():
    package = sample("2010-sample1.docx")
    _ = package.body
    before = part_bytes(package.save())

    with package.dry_run() as trial:
        table = trial.body.insert_table(3, 2, values=[["a", "b"]])
        assert table.row_count == 3
        assert trial.last_change.operation == "insert_table"
        assert trial.last_change.addresses == (table.address,)
        assert len(trial.body.tables) == 1

    assert package.body.tables == []
    assert part_bytes(package.save()) == before


def test_a_dry_run_of_a_picture_leaves_no_part_behind(png):
    package = sample("2010-sample1.docx")
    _ = package.body
    before = part_bytes(package.save(), relationships=True)

    with package.dry_run() as trial:
        picture = trial.body.insert_inline_picture(png, width=20)
        assert picture.image_format == "Png"
        assert trial.last_change.parts_touched[0] == "/word/document.xml"

    assert package.body.inline_pictures == []
    assert part_bytes(package.save(), relationships=True) == before


# ---------------------------------------------------------------------------
# determinism (CR-003 section 3.4)
# ---------------------------------------------------------------------------


def test_the_same_seed_and_the_same_image_give_the_same_part_name_rel_id_and_bytes(png):
    def run() -> tuple[str, str, bytes]:
        package = sample("2010-sample1.docx")
        package.id_seed = SEED
        body = package.body
        body.insert_table(2, 2, values=[["a", "b"], ["c", "d"]], style="TableGrid")
        picture = body.insert_inline_picture(png, width=60, alt_text_description="same")
        return str(picture.image_part.part_name), picture.rel_id, package.save()

    first = run()
    second = run()

    assert first[0] == second[0] == "/word/media/image1.png"
    assert first[1] == second[1]
    assert first[2] == second[2], "the same document and the same calls, the same bytes"


def test_two_pictures_take_the_next_free_name_in_order(png, new_package):
    body = new_package.body
    names = [str(body.insert_inline_picture(png, width=10).image_part.part_name) for _ in range(3)]

    assert names == [
        "/word/media/image1.png",
        "/word/media/image2.png",
        "/word/media/image3.png",
    ]


# ---------------------------------------------------------------------------
# the outline and describe over the views
# ---------------------------------------------------------------------------


def test_the_outline_of_a_document_with_nested_tables():
    package = sample("tables.docx")
    outline = package.outline()
    tables = [entry for entry in outline.entries if entry.kind == "table"]

    assert len(tables) == 13
    for entry in tables:
        view = package.element_at(entry.ordinal)
        assert isinstance(view, Table)
        assert entry.rows == view.row_count
        assert entry.cols == view.rows[0].cell_count
        # the preview is the first row, which is what the view reports too
        first_row = " | ".join(
            value.replace("\n", " ").strip() for value in view.rows[0].values
        )
        assert first_row.startswith(entry.text.rstrip("…"))


def test_describe_still_unmarshals_nothing():
    package = sample("tables.docx")
    before = part_bytes(package.save())

    description = package.describe()
    assert description.styles
    assert description.page.width_pt > 0

    assert part_bytes(package.save()) == before
    for part in package.parts.values():
        if part.part_name.name != "/word/document.xml":
            assert not getattr(part, "is_unmarshalled", False), part.part_name


def test_the_outline_of_the_200_page_document_is_no_slower_than_before():
    """The ``Table`` view went into ``_table_shape``; the outline stays cheap."""
    import time

    package = big_document()
    start = time.perf_counter()
    outline = package.outline()
    elapsed = time.perf_counter() - start

    assert outline.stats.tables == 20
    assert elapsed < 1.0, f"{elapsed:.3f} s"
