"""python-docx code from the wild, run against this engine. CR-003 Phase J.

The other facade test (``tests/content/test_docx_facade.py``) goes member by
member. This one runs **scripts**: the snippet CR-003 section 3.6 promises,
verbatim; the shape of a tutorial script a person or a model writes without
thinking about it; and then the three things python-docx code gets here that it
does not get from python-docx --- a ``ChangeReport`` per call, ``dry_run``, and
every part nobody touched written back byte for byte.
"""

from __future__ import annotations

import io
import zipfile

import pytest
from conftest import SAMPLES, part_bytes

from docx4j_py.docx import WD_ALIGN_PARAGRAPH, Document, Inches, RGBColor

SEED = 20260918


@pytest.fixture(scope="module")
def png() -> bytes:
    """The pangolin from ``samples/Images.docx``."""
    with zipfile.ZipFile(SAMPLES / "Images.docx") as archive:
        name = next(n for n in archive.namelist() if n.lower().endswith(".png"))
        return archive.read(name)


@pytest.fixture
def logo(tmp_path, png) -> str:
    """The pangolin on disk, so a script can name a path as a person would."""
    path = tmp_path / "logo.png"
    path.write_bytes(png)
    return str(path)


# ---------------------------------------------------------------------------
# the snippet CR-003 section 3.6 promises
# ---------------------------------------------------------------------------


def test_the_section_3_6_snippet_runs_verbatim(tmp_path, logo):
    """Exactly the lines in CR-003 section 3.6, paths apart."""
    from docx4j_py.docx import (
        Document,
        Inches,
        Pt,
    )

    out = tmp_path / "out.docx"

    doc = Document()  # or Document("in.docx")
    doc.add_heading("Report", level=1)
    p = doc.add_paragraph("Plain, ")
    p.add_run("bold").bold = True
    p.style = "Heading 2"  # display name, as python-docx
    doc.add_table(rows=2, cols=2).cell(0, 0).text = "a"
    doc.add_picture(logo, width=Inches(2))
    for paragraph in doc.paragraphs:
        assert isinstance(paragraph.style.name, str)
        assert isinstance(paragraph.text, str)
    doc.save(str(out))

    assert Pt(12) == 152400
    again = Document(out)
    assert [p.text for p in again.paragraphs] == ["Report", "Plain, bold", ""]
    assert [p.style.name for p in again.paragraphs] == ["Heading 1", "Heading 2", "Normal"]
    assert again.tables[0].cell(0, 0).text == "a"
    assert again.inline_shapes[0].width == Inches(2)


# ---------------------------------------------------------------------------
# a tutorial script, of the shape a model writes unprompted
# ---------------------------------------------------------------------------

ROWS = [
    {"Region": "North", "Quarter": "Q1", "Revenue": "120"},
    {"Region": "South", "Quarter": "Q1", "Revenue": "98"},
    {"Region": "East", "Quarter": "Q2", "Revenue": "143"},
]


def build_report(logo: str) -> Document:
    """The script: headings, mixed runs, a table from dicts, a picture, a break."""
    document = Document()
    document.core_properties.title = "Quarterly report"
    document.core_properties.author = "Ada Lovelace"

    document.add_heading("Quarterly report", level=0)
    document.add_heading("Summary", level=1)

    paragraph = document.add_paragraph("Revenue rose in ")
    strong = paragraph.add_run("every region")
    strong.bold = True
    paragraph.add_run(" except ")
    weak = paragraph.add_run("the lazy one")
    weak.italic = True
    weak.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    paragraph.add_run(".")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    document.add_heading("By region", level=2)
    headers = list(ROWS[0])
    table = document.add_table(rows=1, cols=len(headers), style="Table Grid")
    for index, name in enumerate(headers):
        cell = table.cell(0, index)
        cell.text = name
        cell.paragraphs[0].runs[0].bold = True
    for row in ROWS:
        cells = table.add_row().cells
        for index, name in enumerate(headers):
            cells[index].text = row[name]

    document.add_page_break()
    document.add_heading("Appendix", level=1)
    document.add_paragraph("The pangolin, for scale:")
    document.add_picture(logo, width=Inches(2))
    document.add_paragraph("Bullet one", style="List Bullet")
    document.add_paragraph("Bullet two", style="List Bullet")
    document.add_paragraph("Step one", style="List Number")
    return document


def test_a_tutorial_script_builds_a_document_that_reads_back(tmp_path, logo):
    document = build_report(logo)
    out = tmp_path / "report.docx"
    document.save(str(out))

    again = Document(out)
    assert [p.style.name for p in again.paragraphs][:4] == [
        "Title",
        "Heading 1",
        "Normal",
        "Heading 2",
    ]
    assert again.paragraphs[2].text == "Revenue rose in every region except the lazy one."
    assert again.paragraphs[2].alignment is WD_ALIGN_PARAGRAPH.JUSTIFY
    assert [r.bold for r in again.paragraphs[2].runs] == [None, True, None, None, None]
    assert [r.italic for r in again.paragraphs[2].runs] == [None, None, None, True, None]
    assert again.paragraphs[2].runs[3].font.color.rgb == RGBColor(0xC0, 0x00, 0x00)

    table = again.tables[0]
    assert len(table.rows) == 4
    assert [c.text for c in table.rows[0].cells] == ["Region", "Quarter", "Revenue"]
    assert [c.text for c in table.rows[3].cells] == ["East", "Q2", "143"]
    assert table.rows[0].cells[0].paragraphs[0].runs[0].bold is True
    assert table.style.name == "Table Grid"

    assert again.inline_shapes[0].width == Inches(2)
    assert again.core_properties.title == "Quarterly report"
    assert [p.style.name for p in again.paragraphs if p.style.name.startswith("List")] == [
        "List Bullet",
        "List Bullet",
        "List Number",
    ]


def test_the_tutorial_script_is_deterministic(logo):
    def bytes_of() -> bytes:
        import datetime as dt

        document = build_report(logo)
        document.package.id_seed = SEED
        fixed = dt.datetime(2026, 9, 18, 10, 0, tzinfo=dt.UTC)
        document.core_properties.created = fixed
        document.core_properties.modified = fixed
        return document.save()

    assert part_bytes(bytes_of()) == part_bytes(bytes_of())


# ---------------------------------------------------------------------------
# a loaded document, edited through the facade
# ---------------------------------------------------------------------------


def test_reading_a_loaded_document_and_saving_leaves_every_other_part_alone():
    """The promise python-docx cannot make: untouched parts are the source's bytes."""
    source = SAMPLES / "2010-sample1.docx"
    original = part_bytes(source.read_bytes())

    document = Document(source)
    texts = [paragraph.text for paragraph in document.paragraphs]
    assert texts and texts[0].startswith("My first")
    saved = part_bytes(document.save())

    assert set(saved) == set(original)
    for name, data in original.items():
        if name == "word/document.xml":
            continue  # reading paragraphs unmarshals this one
        assert saved[name] == data, name


def test_a_loaded_document_saved_unchanged_is_canonically_the_same():
    """CR-002's other rule: an unmarshalled part saved unchanged is equivalent."""
    import sys

    sys.path.insert(0, str(SAMPLES.parent / "scripts"))
    import canon

    source = SAMPLES / "2010-sample1.docx"
    document = Document(source)
    _ = [paragraph.text for paragraph in document.paragraphs]
    saved = part_bytes(document.save())
    before = canon.normalize(part_bytes(source.read_bytes())["word/document.xml"])
    assert canon.normalize(saved["word/document.xml"]) == before


def test_an_edit_through_the_facade_reports_what_it_touched():
    document = Document(SAMPLES / "2010-sample1.docx")
    document.package.changes.clear()

    paragraph = document.paragraphs[0]
    paragraph.add_run(" Edited.").italic = True
    first = document.package.last_change
    assert first.operation.startswith("docx.")
    assert first.parts_touched == ("/word/document.xml",)

    document.add_heading("A new section", level=1)
    styles = document.package.last_change
    assert "/word/document.xml" in styles.parts_touched
    assert "/word/styles.xml" in styles.parts_touched, (
        "Heading 1 is not in this document, so the setter defined it (section 14.9)"
    )

    reports = [change.to_dict() for change in document.package.changes]
    assert all(set(report) >= {"operation", "parts_touched"} for report in reports)


def test_a_dry_run_around_facade_calls_leaves_the_document_alone(png):
    """A trial edits copies; the real document keeps its content and its parts.

    ``add_heading`` makes the trial copy ``/word/styles.xml`` (this document
    defines no Heading 1), and copying a part **unmarshals** the real one, so
    that part is re-serialised on the next save and its bytes change while its
    content does not --- the cost CR-003 section 14.4 names, measured here.
    """
    import sys

    sys.path.insert(0, str(SAMPLES.parent / "scripts"))
    import canon

    document = Document(SAMPLES / "2010-sample1.docx")
    before = part_bytes(document.save())
    paragraph_count = len(document.paragraphs)

    with document.package.dry_run() as trial:
        trial_document = _on_trial(trial)
        trial_document.add_heading("Not kept", level=1)
        trial_document.add_paragraph("Nor this.")
        trial_document.add_table(2, 2, style="Table Grid").cell(0, 0).text = "x"
        trial_document.add_picture(png, width=Inches(1))
        assert len(trial_document.paragraphs) > paragraph_count
        assert trial.last_change is not None

    assert len(document.paragraphs) == paragraph_count
    after = part_bytes(document.save())
    assert set(after) == set(before), "the picture part the trial added went with it"
    assert after["word/document.xml"] == before["word/document.xml"]
    assert "Heading1" not in after["word/styles.xml"].decode(), "no style leaked"
    for name in before:
        if after[name] != before[name]:
            assert name == "word/styles.xml", name
            assert canon.normalize(after[name]) == canon.normalize(before[name])


def _on_trial(trial: object) -> Document:
    """A facade ``Document`` over a trial package, which is all a trial needs."""
    from docx4j_py.docx._base import Story
    from docx4j_py.model.content.body import body_of

    document = Document.__new__(Document)
    document._package = trial
    document._body_story = Story(body_of(trial.get_main_document_part()), document)
    return document


def test_the_content_api_and_the_facade_see_one_document():
    """An edit through either surface is visible through the other."""
    document = Document(SAMPLES / "2010-sample1.docx")
    body = document.part.body

    body.insert_paragraph("From the content API", style="Heading 1")
    assert document.paragraphs[-1].text == "From the content API"
    assert document.paragraphs[-1].style.name == "Heading 1"

    document.add_paragraph("From the facade")
    assert body.paragraphs[-1].text == "From the facade"

    hit = body.search("first")[0]
    hit.font.bold = True
    assert any(run.bold for run in document.paragraphs[0].runs)


def test_a_facade_edit_can_be_tracked_and_reviewed():
    """Tracking is the package's, so it covers facade calls too (Phase F)."""
    document = Document(SAMPLES / "2010-sample1.docx")
    document.package.author = "Claude"
    document.package.change_tracking_mode = "TrackAll"
    document.add_paragraph("Added by an agent.")

    changes = document.package.get_tracked_changes()
    assert changes, "a facade insert is a revision when the mode is on"
    assert {change.author for change in changes} == {"Claude"}

    again = Document(io.BytesIO(document.save()))
    assert again.package.get_tracked_changes()
    # python-docx's `Paragraph.text` is `w:r | w:hyperlink`, so the runs of a
    # tracked insertion are **not** in it --- python-docx's own docstring says
    # "paragraphs within revision marks do not appear" --- and the facade keeps
    # that. The content API reads the accepted view and does see them.
    assert "Added by an agent." not in [p.text for p in again.paragraphs]
    assert "Added by an agent." in again.part.body.text
    assert again.part.body.accept_all() > 0
    assert "Added by an agent." in [p.text for p in again.paragraphs]


def test_a_comment_through_the_facade_is_the_same_comment_phase_g_reads():
    document = Document(SAMPLES / "2010-sample1.docx")
    document.package.changes.clear()
    paragraph = document.paragraphs[0]
    anchor = paragraph.runs[0]
    comment = document.add_comment(anchor, "Please check.", author="Ada", initials="A")

    report = document.package.last_change
    assert "/word/comments.xml" in report.parts_touched
    assert comment.view.author_name == "Ada"
    assert document.part.body.get_comments()[0].content == "Please check."

    again = Document(io.BytesIO(document.save()))
    assert [c.text for c in again.comments] == ["Please check."]
    assert again.part.body.get_comments()[0].author_name == "Ada"
