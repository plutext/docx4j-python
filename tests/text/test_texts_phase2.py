"""The text model's other content, ``docx4j_py.model.text``. CR-011 Phase C (section 13).

docx4j's ``TextsPhase2Test`` at ``198d015a3``, its strings kept: what the
suite's token rows do not pin on their own, built on the probes and on small
documents (``mini_docx.MiniDocx``, docx4j's ``MiniDocx``). Workbooks,
presentations and OLE 1.0 packages are deferred (CR-011 Q7), and their tests
are not ported.
"""

from __future__ import annotations

import pytest
from conformance import PHASE0
from mini_docx import NS, MiniDocx

from docx4j_py import load
from docx4j_py.model.text import ImageDescription, TextResult, TextWarning, extract, plain_text

PROBE = PHASE0 / "extraction-probe.docx"
EXTERNAL = PHASE0 / "extraction-probe-external.docm"


@pytest.fixture(scope="module")
def probe():
    return load(PROBE)


@pytest.fixture(scope="module")
def external():
    return load(EXTERNAL)


def text(package, **options) -> str:
    return plain_text(extract(package, **options))


def between(value: str, start: str, end: str) -> str:
    """The text from `start` up to `end`, `start` included."""
    i = value.index(start)
    return value[i : value.index(end, i + len(start))]


def p(value: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{value}</w:t></w:r></w:p>'


# ---------------------------------------------------------------------------
# metadata (REQ-515)
# ---------------------------------------------------------------------------


def test_metadata_always_in_the_result_and_in_the_text_only_under_metadata(probe, external):
    """REQ-515, REQ-801: the properties always in the metadata; in the text only under metadata: true, before the body."""
    r = extract(external)
    assert r.metadata[TextResult.FIELDS_UPDATE_ON_OPEN] is True
    assert r.metadata[TextResult.APP]["Company"] == "APP52_COMPANY"
    assert r.metadata[TextResult.CUSTOM]["ProbeCustom"] == "CPR52_CUSTOM"
    assert "APP52_COMPANY" not in plain_text(r)
    assert text(external, metadata=True).startswith(
        "[application properties: Application: Microsoft Office Word; Company: APP52_COMPANY]\n"
        "[custom properties: ProbeCustom: CPR52_CUSTOM]\nCase 50"
    )
    assert text(probe, metadata=True).startswith("[core properties: title: PRP38_TITLE; subject: PRP38_SUBJECT]\n")


def test_the_statistics_are_left_out():
    stats = (
        MiniDocx.body(p("Body."))
        .package_part(
            "docProps/app.xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
            "rIdApp",
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
            "<Template>Normal.dotm</Template><Pages>7</Pages><Words>1234</Words><Characters>5678</Characters>"
            "</Properties>",
        )
        .load()
    )
    assert extract(stats).metadata[TextResult.APP] == {"Template": "Normal.dotm"}


def test_every_core_property_in_the_schemas_order_and_custom_ones_typed():
    core = (
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<cp:revision>3</cp:revision><dc:creator>C</dc:creator><dc:title>T</dc:title><cp:keywords>K1 K2</cp:keywords>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-01T00:00:00Z</dcterms:created>'
        "<dc:description></dc:description></cp:coreProperties>"
    )
    custom = (
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Count"><vt:i4>7</vt:i4></property>'
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Done"><vt:bool>true</vt:bool></property>'
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="When">'
        "<vt:filetime>2026-01-02T03:04:05Z</vt:filetime></property>"
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Blob"><vt:clsid>{00000000-0000-0000-0000-000000000000}</vt:clsid></property>'
        "</Properties>"
    )
    pkg = (
        MiniDocx.body(p("Body."))
        .package_part(
            "docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "rIdCore",
            core,
        )
        .package_part(
            "docProps/custom.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
            "rIdCustom",
            custom,
        )
        .load()
    )
    r = extract(pkg, metadata=True)
    assert list(r.metadata[TextResult.CORE].items()) == [
        ("title", "T"),
        ("creator", "C"),
        ("keywords", "K1 K2"),
        ("revision", "3"),
        ("created", "2026-10-01T00:00:00Z"),
    ]
    assert r.metadata[TextResult.CUSTOM] == {"Count": 7, "Done": True, "When": "2026-01-02T03:04:05Z"}
    assert plain_text(r).startswith(
        "[core properties: title: T; creator: C; keywords: K1 K2; revision: 3; created: 2026-10-01T00:00:00Z]\n"
        "[custom properties: Count: 7; Done: true; When: 2026-01-02T03:04:05Z]\nBody.\n"
    )
    assert r.to_dict()["metadata"]["custom"]["Done"] is True


# ---------------------------------------------------------------------------
# images and objects (REQ-514, REQ-512)
# ---------------------------------------------------------------------------


def runs(result, kind):
    return [r for b in result.blocks for r in b.runs if r.kind == kind]


def test_image_descriptions(probe):
    """REQ-514, REQ-707: the alt text, else the name; nothing for a decorative picture; the title never."""
    value = text(probe, alt_text=True)
    assert between(value, "Alt text and title:", "Case 34") == (
        "Alt text and title: [image: IMG33_ALT_TEXT]\nDecorative: \nSVG, its PNG fallback in a:blip: [image: IMG33_SVG_ALT]\n"
    )
    assert "IMG33_TITLE" not in value
    assert "[image" not in text(probe)
    images = runs(extract(probe), "IMAGE")
    decorative = [r for r in images if r.image.decorative]
    assert len(decorative) == 1
    assert decorative[0].description is None
    assert decorative[0].image.descr == "IMG33_DECORATIVE_DESCR"
    assert all(r.description.provenance == "AUTHOR" for r in images if r.description is not None)


def test_a_hosts_describer_is_marked_generated(probe):
    def captions(image):
        return None if image.decorative else ImageDescription.generated(f"a picture named {image.name}")

    assert "Alt text and title: [image, generated: a picture named Picture 14]" in text(
        probe, alt_text=True, image_describer=captions
    )


def test_a_vml_pictures_alt_and_no_decorative_block_in_the_probe(probe):
    vml = MiniDocx.body(
        '<w:p><w:r><w:pict><v:shape style="width:10pt;height:10pt" alt="VML alt"><v:imagedata r:id="rIdNone"/>'
        "</v:shape></w:pict></w:r></w:p>"
    ).load()
    assert text(vml, alt_text=True) == "[image: VML alt]\n"
    assert not any(b.decorative for b in extract(probe).blocks)


def test_a_paragraph_of_decorative_pictures_is_a_decorative_block():
    picture = (
        '<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="P" descr="D">'
        '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">'
        '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>'
        "</a:ext></a:extLst></wp:docPr><a:graphic>"
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
        '<pic:cNvPr id="0" name="P"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdNone"/>'
        "</pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )
    block = extract(MiniDocx.body(f"<w:p>{picture}</w:p>").load()).blocks[0]
    assert block.decorative and block.to_dict()["decorative"]
    assert block.runs[0].image.part is None  # the relationship names nothing


def test_a_describer_that_fails_raises_image_not_described(probe):
    def failing(image):
        raise RuntimeError("the describing service is down")

    r = extract(probe, alt_text=True, image_describer=failing)
    found = [w for w in r.warnings if w.kind == TextWarning.IMAGE_NOT_DESCRIBED]
    assert found and all(w.selector.startswith("/word/media/") for w in found)
    assert "[image" not in plain_text(r)


def test_objects_are_object_runs_with_their_progids(probe):
    assert [r.prog_id for r in runs(extract(probe), "OBJECT")] == [
        "Excel.Sheet.12",
        "Word.Document.12",
        "Package",
        None,
        "PowerPoint.Show.12",
    ]


def test_legacy_wordart_is_a_one_paragraph_text_box(probe):
    r = extract(probe)
    word_art = [b for b in r.blocks if b.story == "TEXT_BOX" and b.text == "WA25_TEXTPATH"]
    assert len(word_art) == 1 and word_art[0].address.endswith("/textbox:0/0")
    assert "WA25_TEXTPATH" not in text(probe, text_boxes=False)


# ---------------------------------------------------------------------------
# tables (REQ-605)
# ---------------------------------------------------------------------------


def tc(props: str, *paragraphs: str) -> str:
    cells = "".join("<w:p/>" if not value else p(value) for value in paragraphs)
    return f"<w:tc><w:tcPr>{props}</w:tcPr>{cells}</w:tc>"


def grid(*widths: int) -> str:
    return "<w:tbl><w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths) + "</w:tblGrid>"


SPAN2, SPAN3 = '<w:gridSpan w:val="2"/>', '<w:gridSpan w:val="3"/>'
HEADER = "<w:trPr><w:tblHeader/></w:trPr>"


def test_table_forms(probe):
    """REQ-605: case 31 in the listed format, as TSV under tables: tsv, as a grid under tables: grid."""
    listed = text(probe)
    # its one header cell spans the grid: written once, labelling nothing, the rows' cells a tab apart
    assert between(listed, "[table: listed format]\nTBL31", "After the table.\nCase 32") == (
        "[table: listed format]\nTBL31_HEADER\nTBL31_CELL_P1; TBL31_CELL_P2\tTBL31_NESTED\n"
        "TBL31_MERGED\tTBL31_LAST_CELL\n"
    )
    # a block-level control in a cell, with Word's empty paragraph after it, is a simple table's cell
    assert between(listed, "SDT24_IN_CELL", "After the table.\nCase 25") == "SDT24_IN_CELL\n"
    assert between(text(probe, tables="tsv"), "TBL31_HEADER", "After the table.\nCase 32") == (
        "TBL31_HEADER\nTBL31_CELL_P1 TBL31_CELL_P2\tTBL31_NESTED\t\nTBL31_MERGED\tTBL31_LAST_CELL\n"
    )
    assert between(text(probe, tables="grid"), "Case 31: a table", "After the table.\nCase 32") == (
        "Case 31: a table: header row, two-paragraph cell, nested table, merged cell, empty cell\n"
        "+---------------+------------------+-----------------+\n"
        "| TBL31_HEADER                                       |\n"
        "+===============+==================+=================+\n"
        "| TBL31_CELL_P1 | +--------------+ |                 |\n"
        "|               | | TBL31_NESTED | |                 |\n"
        "| TBL31_CELL_P2 | +--------------+ |                 |\n"
        "+---------------+------------------+-----------------+\n"
        "| TBL31_MERGED                     | TBL31_LAST_CELL |\n"
        "+---------------+------------------+-----------------+\n"
    )


MERGED_TABLE = (
    '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
    f"<w:tr>{tc('', 'Name')}{tc('', 'Value')}</w:tr>"
    f"<w:tr>{tc('<w:vMerge w:val=\"restart\"/>', 'a')}{tc('', '1')}</w:tr>"
    f"<w:tr>{tc('<w:vMerge/>', '')}{tc('', '2')}</w:tr>"
    f"<w:tr>{tc('', '')}{tc('', '3')}</w:tr></w:tbl>{p('After.')}"
)


def test_listed_format_and_grid_spans():
    """REQ-605: headers from the first row; a vertical merge's value on each row it covers; an empty cell left out."""
    pkg = MiniDocx.body(MERGED_TABLE).load()
    assert text(pkg) == "[table: listed format]\nName: a; Value: 1\nName: a; Value: 2\nValue: 3\nAfter.\n"
    assert text(pkg, tables="grid") == (
        "+------+-------+\n"
        "| Name | Value |\n"
        "+======+=======+\n"
        "| a    | 1     |\n"
        "|      +-------+\n"
        "|      | 2     |\n"
        "+------+-------+\n"
        "|      | 3     |\n"
        "+------+-------+\n"
        "After.\n"
    )
    # a row whose one cell spans the grid is its text alone, a header row (labelling nothing) or a body row
    headed = MiniDocx.body(
        '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
        f"<w:tr><w:trPr><w:tblHeader/></w:trPr>{tc('', 'H1')}{tc('', 'H2')}</w:tr>"
        f"<w:tr>{tc(SPAN2, 'wide')}</w:tr></w:tbl>"
        '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
        f"<w:tr>{tc(SPAN2, 'only')}</w:tr><w:tr>{tc('', 'x')}{tc('', 'y')}</w:tr></w:tbl>"
    ).load()
    assert text(headed) == "[table: listed format]\nwide\n[table: listed format]\nonly\nx\ty\n"


def test_listed_format_headers():
    """REQ-605's listed format, as ruled on 18 October 2026 (docx4j's listedFormatHeaders)."""
    rows = f"<w:tr>{tc('', 'P1', 'P2')}{tc('', 'N')}{tc('', '')}</w:tr><w:tr>{tc(SPAN2, 'M')}{tc('', 'L')}</w:tr></w:tbl>"

    def of(body: str) -> str:
        return text(MiniDocx.body(body).load())

    # one header a column; a cell spanning two headed columns takes the first's
    assert of(grid(3000, 3000, 3000) + f"<w:tr>{HEADER}{tc('', 'A')}{tc('', 'B')}{tc('', 'C')}</w:tr>" + rows) == (
        "[table: listed format]\nA: P1; P2; B: N\nA: M; C: L\n"
    )
    # a header spanning two columns labels each cell under it
    assert of(grid(3000, 3000, 3000) + f"<w:tr>{HEADER}{tc(SPAN2, 'AB')}{tc('', 'C')}</w:tr>" + rows) == (
        "[table: listed format]\nAB: P1; P2; AB: N\nAB: M; C: L\n"
    )
    # a group header over its sub-headers; a header with no text over the first column
    assert of(
        grid(3000, 3000, 3000)
        + f"<w:tr>{HEADER}{tc('<w:vMerge w:val=\"restart\"/>', 'Region')}{tc(SPAN2, 'Sales')}</w:tr>"
        + f"<w:tr>{HEADER}{tc('<w:vMerge/>', '')}{tc('', '2024')}{tc('', '2025')}</w:tr>"
        + f"<w:tr>{tc('', 'North')}{tc('', '1')}{tc('', '2')}</w:tr></w:tbl>"
    ) == "[table: listed format]\nRegion: North; Sales / 2024: 1; Sales / 2025: 2\n"
    assert of(
        grid(3000, 3000, 3000)
        + f"<w:tr>{HEADER}{tc('', '')}{tc('', 'B')}{tc('', 'C')}</w:tr>"
        + f"<w:tr>{tc(SPAN2, 'ab')}{tc('', 'c')}</w:tr>"
        + f"<w:tr>{tc('', 'x')}{tc('', 'y')}{tc('', 'z')}</w:tr></w:tbl>"
    ) == "[table: listed format]\nB: ab; C: c\nx; B: y; C: z\n"
    # a title row over column headers, and a divider row in the body
    assert of(
        grid(3000, 3000)
        + f"<w:tr>{HEADER}{tc(SPAN2, 'Title')}</w:tr>"
        + f"<w:tr>{HEADER}{tc('', 'H1')}{tc('', 'H2')}</w:tr>"
        + f"<w:tr>{tc('', 'a')}{tc('', 'b')}</w:tr>"
        + f"<w:tr>{tc(SPAN2, 'Part B')}</w:tr>"
        + f"<w:tr>{tc('', 'c')}{tc('', 'd')}</w:tr></w:tbl>"
    ) == "[table: listed format]\nTitle\nH1: a; H2: b\nPart B\nH1: c; H2: d\n"
    # every header row spans the grid: the body is unlabelled (no data row promoted), its cells a tab apart
    assert of(
        grid(3000, 3000, 3000)
        + f"<w:tr>{HEADER}{tc(SPAN3, 'Title')}</w:tr>"
        + f"<w:tr>{tc('', 'a', 'a2')}{tc('', '')}{tc('', 'b')}</w:tr></w:tbl>"
    ) == "[table: listed format]\nTitle\na; a2\tb\n"


# ---------------------------------------------------------------------------
# page markers and indents (REQ-708, REQ-709)
# ---------------------------------------------------------------------------


def test_page_markers():
    """REQ-708: [page n] at each w:lastRenderedPageBreak of the main story, under page_markers: true only."""
    resaved = load(PHASE0.parent / "extraction-probe.word-resave.docx")
    assert "[page " not in text(resaved)
    value = text(resaved, page_markers=True)
    markers = runs(extract(resaved), "PAGE_MARKER")
    assert markers
    for i, marker in enumerate(markers):
        assert marker.page_number == i + 2
        assert f"[page {i + 2}]" in value
    assert f"[page {len(markers) + 2}]" not in value


def test_layout_indent():
    """REQ-709: four spaces a 720-twip tab stop of a paragraph's first-line indent; tabs kept; none in a table."""
    pkg = (
        MiniDocx.body(
            '<w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>one stop</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:ind w:left="1440" w:firstLine="720"/></w:pPr><w:r><w:t>three</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:ind w:left="1440" w:hanging="720"/></w:pPr><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>'
            + p("none")
            + '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:pPr><w:ind w:left="720"/>'
            "</w:pPr><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        )
        .part(
            "styles.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "rIdStyles",
            f"<w:styles {NS}/>",
        )
        .load()
    )
    assert text(pkg) == "one stop\nthree\na\tb\nnone\ncell\n"
    assert text(pkg, layout="indent") == "    one stop\n            three\n    a\tb\nnone\ncell\n"


# ---------------------------------------------------------------------------
# math (REQ-604)
# ---------------------------------------------------------------------------

EQUATIONS = PHASE0.parent.parent / "revisions" / "tracked-changes-equations.docx"


def test_math_as_latex(probe):
    """REQ-604: $$...$$ for an m:oMathPara, $...$ inline; in markup an equation's tracked changes as its two forms."""
    assert between(text(probe), "$$MTH32", "Case 33") == "$$MTH32\\_A+\\frac{MTH32\\_B}{2}$$\n"
    inline = MiniDocx.body(
        '<w:p><w:r><w:t xml:space="preserve">So </w:t></w:r><m:oMath><m:r><m:t>x</m:t></m:r>'
        "<m:sSup><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></w:p>"
    ).load()
    assert text(inline) == "So $xy^{2}$\n"


@pytest.mark.parametrize("view", ["accepted", "original"])
def test_the_walk_and_the_resolved_copy_agree_on_equations(view):
    from docx4j_py.model import revisions
    from docx4j_py.model.text import extract_resolved

    copy = load(EQUATIONS)
    (revisions.accept_all if view == "accepted" else revisions.reject_all)(copy)
    assert plain_text(extract_resolved(copy, view=view)) == text(load(EQUATIONS), view=view)


def test_a_structure_inserted_whole_is_in_the_accepted_form_only():
    tracked = load(EQUATIONS)
    line = "with units of metres per second:\n"
    after = between(text(tracked), line, "\nHeight")[len(line) :]
    before = between(text(tracked, view="original"), line, "\nHeight")[len(line) :]
    assert after.startswith("$$\\mathrm{m}\\mathrm{") and after.endswith("}{\\mathrm{s}}^{\\mathrm{-1}}\\mathrm{.}$$")
    assert "^{\\mathrm" not in before
    assert between(text(tracked, view="markup"), line, "\nHeight") == f"{line}{{--{before}--}}{{++{after}++}}"


def test_an_equation_outside_the_subset_is_its_text():
    """An m:func is not in docx4j's subset: the runs' text, delimited."""
    pkg = MiniDocx.body(
        "<w:p><m:oMath><m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e>"
        "</m:func></m:oMath></w:p>"
    ).load()
    assert text(pkg) == "$sinx$\n"


@pytest.mark.parametrize(
    ("omml", "latex"),
    [
        ("<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>sin</m:t></m:r><m:r><m:t>x</m:t></m:r>", "\\sin x"),
        ("<m:r><m:t>a≤b</m:t></m:r>", "a\\le b"),
        ("<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/><m:e><m:r><m:t>2</m:t></m:r></m:e></m:rad>", "\\sqrt{2}"),
        ("<m:nary><m:naryPr><m:chr m:val=\"∑\"/></m:naryPr><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup/>"
         "<m:e><m:r><m:t>i</m:t></m:r></m:e></m:nary>", "\\sum_{i}^{}i"),
        ("<m:d><m:dPr><m:begChr m:val=\"[\"/><m:endChr m:val=\"]\"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>",
         "\\left[x\\right]"),
        ("<m:acc><m:accPr><m:chr m:val=\"̃\"/></m:accPr><m:e><m:r><m:t>a</m:t></m:r></m:e></m:acc>", "\\tilde{a}"),
        ("<m:r><m:rPr><m:nor/></m:rPr><m:t>if x_1</m:t></m:r>", "\\text{if x\\_1}"),
    ],
)  # fmt: skip
def test_omml_to_latex(omml, latex):
    from docx4j_py.model.math import OmmlToLatex
    from docx4j_py.wml import wml

    p = wml(f'<w:p {NS}><m:oMath>{omml}</m:oMath></w:p>')
    from docx4j_py.model.revisions import trees as rtrees

    equation = next(item for item in p.content if rtrees.is_math(item))
    assert OmmlToLatex().convert(equation) == latex


# ---------------------------------------------------------------------------
# bound values (REQ-509)
# ---------------------------------------------------------------------------

STORE = "{11111111-2222-3333-4444-555555555555}"


def bound(body: str):
    return (
        MiniDocx.body(body)
        .relationship(
            "rIdXml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml", "../customXml/item1.xml"
        )
        .entry("customXml/item1.xml", None, '<d xmlns="urn:t"><a>Bound A</a><empty/></d>')
        .entry(
            "customXml/_rels/item1.xml.rels",
            None,
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" '
            'Target="itemProps1.xml"/></Relationships>',
        )
        .entry(
            "customXml/itemProps1.xml",
            "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
            f'<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{STORE}"/>',
        )
        .package_part(
            "docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "rIdCore",
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Core title</dc:title></cp:coreProperties>',
        )
        .load()
    )


def control(kind: str, binding: str, placeholder: bool, surface: str) -> str:
    return (
        f"<w:p><w:sdt><w:sdtPr>{'<w:showingPlcHdr/>' if placeholder else ''}{binding}{kind}</w:sdtPr><w:sdtContent>"
        f'<w:r><w:t xml:space="preserve">{surface}</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    )


def binding(xpath: str, store: str, mappings: str = "xmlns:t='urn:t'") -> str:
    return f'<w:dataBinding w:prefixMappings="{mappings}" w:xpath="{xpath}" w:storeItemID="{store}"/>'


def test_bound_values(probe):
    """REQ-509: the bound value where the binding resolves; else the surface; the placeholder where the value is empty."""
    r = extract(probe)
    assert "The control: BND23_XML_VALUE\n" in plain_text(r)
    block = next(b for b in r.blocks if "BND23_XML_VALUE" in b.text)
    assert TextWarning(TextWarning.SDT_BINDING_DIFFERS, block.address) in r.warnings
    surface = extract(probe, bindings="surface")
    assert "The control: BND23_SURFACE_STALE\n" in plain_text(surface)
    assert any(w.kind == TextWarning.SDT_BINDING_DIFFERS for w in surface.warnings)

    core = "xmlns:ns0='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:ns1='http://purl.org/dc/elements/1.1/'"
    pkg = bound(
        control("<w:text/>", binding("/t:d/t:a", STORE), False, "Bound A")  # agrees: no warning
        + control("<w:text/>", binding("/t:d/t:empty", STORE), True, "Click here")  # empty, a placeholder shown
        + control("<w:text/>", binding("/t:d/t:missing", STORE), False, "Surface C")  # doesn't resolve
        + control("", binding("/t:d/t:a", STORE), False, "Rich D")  # rich text: the surface
        + control(
            "<w:text/>",
            binding("/ns0:coreProperties[1]/ns1:title[1]", "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", core),
            False,
            "Old title",
        )  # the core properties
    )
    before = pkg.save()
    b = extract(pkg)
    assert plain_text(b) == "Bound A\nClick here\nSurface C\nRich D\nCore title\n"
    assert b.warnings == [TextWarning(TextWarning.SDT_BINDING_DIFFERS, "body/4")]
    assert pkg.save() == before  # REQ-204


# ---------------------------------------------------------------------------
# charts, SmartArt (REQ-511)
# ---------------------------------------------------------------------------


def test_charts_and_smartart(probe):
    """REQ-511: a chart's title and table as a block after its paragraph; SmartArt's lines in reading order."""
    r = extract(probe)
    chart = next(b for b in r.blocks if b.kind == "CHART")
    assert chart.address.endswith("/chart:0") and chart.story == "MAIN"
    assert chart.chart_title == "CHT34_TITLE"
    # its description's fields (CR-047, REQ-706), and its note before the title in plain text only (REQ-707)
    assert chart.chart.kind == "clustered column"
    assert chart.chart.axes[1].title == "CHT34_AXIS_TITLE"
    assert "[chart" not in chart.text
    assert between(plain_text(r), "[chart:", "Case 35") == (
        "[chart: clustered column; value axis: CHT34_AXIS_TITLE]\nCHT34_TITLE\n\tCHT34_SERIES_ONE\tCHT34_SERIES_TWO\n"
        "CHT34_CAT_A\t3\t4\nCHT34_CAT_B\t5\t2\n"
    )
    assert chart.chart_header == ["", "CHT34_SERIES_ONE", "CHT34_SERIES_TWO"]
    assert chart.chart_rows == [["CHT34_CAT_A", "3", "4"], ["CHT34_CAT_B", "5", "2"]]
    data = chart.to_dict()
    assert data["chart_kind"] == "clustered column" and data["secondary_series"] == []
    assert data["axes"] == [{"role": "category axis", "title": None}, {"role": "value axis", "title": "CHT34_AXIS_TITLE"}]
    assert not [b for b in extract(probe, charts=False).blocks if b.kind == "CHART"]

    smart_art = [b for b in r.blocks if b.story == "SMARTART"]
    assert [b.text for b in smart_art] == ["SMA35_ROOT", "SMA35_CHILD_ONE", "SMA35_CHILD_TWO"]
    assert smart_art[2].address.endswith("/smartart:0/2")
    assert not [b for b in extract(probe, smartart=False).blocks if b.story == "SMARTART"]


def test_a_chart_or_smartart_that_cannot_be_read_raises_content_not_read():
    """REQ-801: content that can't be read contributes nothing; the part's name its selector, the anchor's address its address."""
    inline = (
        '<w:p><w:r><w:t>Chart</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="C"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC"/>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
        '<w:p><w:r><w:t>Diagram</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="2" name="D"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdD" r:lo="" r:qs="" r:cs=""/>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )
    drawings = (
        MiniDocx.body(inline)
        .part(
            "charts/chart1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            "rIdC",
            "<notAChart/>",
        )
        .part(
            "diagrams/data1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
            "rIdD",
            "<notData/>",
        )
        .load()
    )
    r = extract(drawings)
    assert TextWarning(TextWarning.CONTENT_NOT_READ, "body/0", part_name="/word/charts/chart1.xml") in r.warnings
    assert TextWarning(TextWarning.CONTENT_NOT_READ, "body/1", part_name="/word/diagrams/data1.xml") in r.warnings
    assert plain_text(r) == "Chart\nDiagram\n"


def test_a_chartex_chart_in_its_choice():
    """waterfall.docx's chart, a cx:chart in an mc:Choice (Requires cx1, understood), is read; its picture is not."""
    r = extract(load(PHASE0.parent.parent / "waterfall.docx"))
    charts = [b for b in r.blocks if b.kind == "CHART"]
    assert len(charts) == 1 and charts[0].chart.note == "[chart: waterfall]"
    assert charts[0].chart_header == ["", "Series1"] and charts[0].chart_rows[3] == ["Category 4", "-40"]


# ---------------------------------------------------------------------------
# the glossary (REQ-501, REQ-516)
# ---------------------------------------------------------------------------


def test_the_glossary_last(probe):
    """REQ-501, REQ-516: the glossary only under include: glossary, last, after the endnotes."""
    assert "GLS54" not in text(probe)
    r = extract(probe, include="glossary")
    glossary = [b for b in r.blocks if b.story == "GLOSSARY"]
    assert len(glossary) == 2
    assert glossary[0].address == "glossary:0/0"
    assert glossary[1].text == "GLS54_ENTRY_TWO"
    assert r.blocks[-2:] == glossary
    assert plain_text(r).endswith("[1] EN27_NOTE_TEXT\nGLS54_ENTRY_ONE\nGLS54_ENTRY_TWO\n")


# ---------------------------------------------------------------------------
# altChunk (REQ-513)
# ---------------------------------------------------------------------------

A_F_CHUNK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"
HTML = (
    "<html><head><title>T</title><style>p{}</style></head><body><p>One &amp; two</p>"
    "<div>Three<br>Four</div><table><tr><td>a</td><td>b</td></tr></table><script>no()</script></body></html>"
)


def test_alt_chunks():
    """REQ-513: text chunks as lines; HTML not expanded without an HTML reader (CR-011 Q13); RTF never; one level."""
    pkg = (
        MiniDocx.body(f'{p("Before.")}<w:altChunk r:id="rIdT"/><w:altChunk r:id="rIdR"/><w:altChunk r:id="rIdH"/>{p("After.")}')
        .part("t.txt", "text/plain", A_F_CHUNK, "rIdT", b"line one\r\nline two\r\n")
        .part("r.rtf", "application/rtf", A_F_CHUNK, "rIdR", b"{\\rtf1 RTF text}")
        .part("h.htm", "text/html", A_F_CHUNK, "rIdH", HTML.encode())
        .load()
    )
    before = pkg.save()
    r = extract(pkg)
    assert plain_text(r) == "Before.\nline one\nline two\nAfter.\n"
    assert [str(w) for w in r.warnings] == [
        "altchunk.not_expanded at body/2 (application/rtf)",
        "altchunk.not_expanded at body/3 (text/html)",
    ]
    assert r.blocks[2].address == "body/1/altchunk:0/1"
    off = extract(pkg, altchunks=False)
    assert plain_text(off) == "Before.\nAfter.\n"
    assert len(off.warnings) == 3 and off.warnings[0].content_type == "text/plain"
    assert pkg.save() == before  # REQ-204


def test_a_document_chunks_html_chunk_is_not_expanded(probe):
    """The probe's document chunk holds an HTML chunk, which is not expanded (one level)."""
    nested = [w for w in extract(probe).warnings if w.address is not None and "/altchunk:0/" in w.address]
    assert len(nested) == 1 and nested[0].content_type == "text/html"


def test_a_chunks_text_is_decoded_by_its_bom_else_utf8_else_windows_1252():
    from docx4j_py.model.text.embedded import decode

    assert decode("é".encode()) == "é"
    assert decode(b"\xef\xbb\xbf" + "é".encode()) == "é"
    assert decode(b"\xff\xfe" + "é".encode("utf-16-le")) == "é"
    assert decode(b"\xe9") == "é"


# ---------------------------------------------------------------------------
# embedded packages (REQ-512)
# ---------------------------------------------------------------------------


def test_an_embedded_document_is_read_one_level_deep(probe):
    """REQ-512: under embedded: true, a WordprocessingML package read at its reference, one level deep."""
    r = extract(probe, embedded=True)
    texts = {b.address.split("/", 2)[2]: b.text for b in r.blocks if b.story == "EMBEDDED"}
    assert texts["object:0/body/0"] == "EMB40_DOCX_TEXT"
    value = plain_text(r)
    assert "EMB40_NESTED_TEXT" not in value and "ACX42_CAPTION" not in value
    assert "EMB40_DOCX_TEXT" not in text(probe)


def test_the_deferred_packages_raise_content_not_read(probe):
    """CR-011 Q7: a workbook, a presentation and an OLE 1.0 package contribute nothing, and say so."""
    r = extract(probe, embedded=True)
    not_read = {w.selector for w in r.warnings if w.kind == TextWarning.CONTENT_NOT_READ}
    assert not_read == {"/word/embeddings/probe39.xlsx", "/word/embeddings/probe41.bin", "/word/embeddings/probe53.pptx"}
    assert not [w for w in extract(probe).warnings if w.kind == TextWarning.CONTENT_NOT_READ]


def test_an_embedded_package_that_is_not_one_raises_content_not_read():
    obj = (
        MiniDocx.body(
            '<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:object><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" '
            'r:id="rIdE"/></w:object></w:r></w:p>'
        )
        .part(
            "embeddings/broken.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
            "rIdE",
            b"not a zip",
        )
        .load()
    )
    r = extract(obj, embedded=True)
    assert TextWarning(TextWarning.CONTENT_NOT_READ, "body/0", part_name="/word/embeddings/broken.xlsx") in r.warnings
    assert plain_text(r) == "Before\n"
    assert not extract(obj).warnings  # not read at all without embedded: true


def test_an_embedded_documents_hidden_text_under_hidden():
    inspect = load(PHASE0 / "extraction-probe-inspect.docx")
    assert "EMB63_VISIBLE" in text(inspect, embedded=True) and "EMB63_HIDDEN" not in text(inspect, embedded=True)
    assert "EMB63_HIDDEN" in text(inspect, embedded=True, hidden=True)


# ---------------------------------------------------------------------------
# fields: update (REQ-506)
# ---------------------------------------------------------------------------


def complex_field(instr: str, cached: str, locked: bool = False) -> str:
    return (
        f'<w:r><w:fldChar w:fldCharType="begin"{" w:fldLock=\"1\"" if locked else ""}/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
        f'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{cached}</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def fields_package():
    body = (
        f"<w:p>{complex_field('DOCPROPERTY Client', 'old client')}</w:p>"
        '<w:p><w:fldSimple w:instr=" TITLE \\* Upper "><w:r><w:t>old title</w:t></w:r></w:fldSimple></w:p>'
        f"<w:p>{complex_field('DOCVARIABLE v', 'old var')}</w:p>"
        f"<w:p>{complex_field('SEQ Figure', '9')}{complex_field('SEQ Figure', '9')}{complex_field('SEQ Figure \\r 7', '9')}"
        f"{complex_field('SEQ Figure \\c', '9')}</w:p>"
        f"<w:p>{complex_field('REF later \\h', 'old ref')}</w:p>"
        f"<w:p>{complex_field('QUOTE \"quoted words\"', 'old quote')}</w:p>"
        f"<w:p>{complex_field('DATE \\@ \"d MMMM yyyy\"', '1 January 2000')}</w:p>"
        f"<w:p>{complex_field('DOCPROPERTY Client', 'locked', locked=True)}</w:p>"
        '<w:p><w:bookmarkStart w:id="0" w:name="later"/><w:r><w:t>Target text</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>'
    )
    return (
        MiniDocx.body(body)
        .part(
            "settings.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
            "rIdSettings",
            f'<w:settings {NS}><w:docVars><w:docVar w:name="v" w:val="Var value"/></w:docVars></w:settings>',
        )
        .package_part(
            "docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "rIdCore",
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Probe title</dc:title></cp:coreProperties>',
        )
        .package_part(
            "docProps/custom.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
            "rIdCustom",
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Client"><vt:lpwstr>Acme</vt:lpwstr>'
            "</property></Properties>",
        )
        .load()
    )


def test_fields_update():
    """REQ-506: cached results by default; under fields: update, the fields that need neither layout nor fetching recomputed."""
    pkg = fields_package()
    assert text(pkg) == "old client\nold title\nold var\n9999\nold ref\nold quote\n1 January 2000\nlocked\nTarget text\n"
    before = pkg.save()
    r = extract(pkg, fields="update")
    assert plain_text(r) == "Acme\nPROBE TITLE\nVar value\n1277\nTarget text\nquoted words\n1 January 2000\nlocked\nTarget text\n"
    assert r.blocks[0].fields[0].updated and r.blocks[0].fields[0].result == "Acme"
    assert not r.blocks[6].fields[0].updated
    assert not r.blocks[7].fields[0].updated
    assert pkg.save() == before  # REQ-204


def test_a_computed_fields_cached_result_is_not_read():
    """REQ-506: a REF whose cached result holds a note reference and a nested field writes the bookmark's text alone."""
    pkg = (
        MiniDocx.body(
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> REF b \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>Old</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:fldSimple w:instr=" DOCPROPERTY Client "/>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
            '<w:p><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>Target</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>'
            '<w:bookmarkEnd w:id="0"/></w:p>'
        )
        .footnotes(
            '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> NOTE_ONE</w:t></w:r>'
            "</w:p></w:footnote>"
        )
        .load()
    )
    assert text(pkg) == "Old[1]\nTarget[1]\n[1] NOTE_ONE\n"
    r = extract(pkg, fields="update")
    assert plain_text(r) == "Target\nTarget[1]\n[1] NOTE_ONE\n"
    first = r.blocks[0]
    assert all(run.kind != "NOTE_REFERENCE" for run in first.runs)
    assert len(first.fields) == 1 and first.fields[0].updated
    assert not [w for w in r.warnings if w.kind == TextWarning.FIELD_NO_RESULT]
    codes = extract(pkg, fields="update", field_codes=True)
    assert "".join(run.text for run in codes.blocks[0].runs if run.kind == "FIELD_CODE") == " REF b \\h "


def test_ref_on_the_part_route():
    """REQ-506 on a part's route, which has no first reading: a REF after its bookmark computed, one before it cached."""
    ref = (
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF b \\h </w:instrText>'
        '</w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Old</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/>'
        "</w:r></w:p>"
    )
    bookmark = '<w:p><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>Target</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>'
    after = MiniDocx.body(bookmark + ref).load()
    assert plain_text(extract(after.main_document_part, fields="update")) == "Target\nTarget\n"
    before = MiniDocx.body(ref + bookmark).load()
    assert plain_text(extract(before.main_document_part, fields="update")) == "Old\nTarget\n"
    assert plain_text(extract(before, fields="update")) == "Target\nTarget\n"


def test_the_main_parts_route_reads_its_story_alone(probe):
    """docx4j ``Texts.of(Part)``: the main story and its text boxes, addressed body; no notes; no first reading."""
    r = extract(probe.main_document_part)
    assert r.blocks[0].address == "body/0" and all(b.story in ("MAIN", "TEXT_BOX", "SMARTART") for b in r.blocks)
    assert "FN27_NOTE_TEXT" not in plain_text(r)
    header = next(iter(probe.main_document_part.header_parts()))
    h = extract(header)
    assert h.blocks and h.blocks[0].address.startswith("header:rId") and h.blocks[0].story == "HEADER"


FORMTEXT = PHASE0.parent / "formtext"


def test_formtext_as_word_updates_it():
    """REQ-506: FORMTEXT under fields: update as Word's update leaves it, Word's own save after F9 the oracle."""
    before = load(FORMTEXT / "cr046-formtext-unprotected.docx")
    word = load(FORMTEXT / "cr046-formtext-unprotected-after-f9.docx")
    ours = text(before, fields="update")
    words = text(word)
    a, b = "A: typed over a default: FT_DEFAULT\n", "B: typed, no default:      \n"
    assert words.startswith(a + b)
    assert ours.startswith(a + b)  # C, a formula, is not computed
    typed = "A: typed over a default: FT_TYPED\nB: typed, no default: FT_TYPED_NO_DEFAULT\n"
    assert text(before).startswith(typed)
    protected = load(FORMTEXT / "cr046-formtext-protected-update-on-open.docx")
    assert text(protected, fields="update").startswith(typed)


# ---------------------------------------------------------------------------
# REQ-204: with every option on, the package is as it was
# ---------------------------------------------------------------------------

EVERYTHING = {
    "fields": "update",
    "embedded": True,
    "metadata": True,
    "alt_text": True,
    "include": "headers, footers, comments, glossary",
    "tables": "grid",
    "page_markers": True,
    "layout": "indent",
    "hidden": True,
    "list_labels": True,
    "field_codes": True,
}


@pytest.mark.parametrize(
    "path",
    [PROBE, EXTERNAL, PHASE0 / "extraction-probe-structure.docx", PHASE0 / "extraction-probe-inspect.docx"],
    ids=lambda p: p.name,
)
def test_with_every_option_the_package_is_unchanged(path):
    package = load(path)
    before = package.save()
    for view in ("accepted", "original", "markup"):
        extract(package, view=view, **EVERYTHING)
    assert not package.main_document_part.is_unmarshalled
    assert package.save() == before
