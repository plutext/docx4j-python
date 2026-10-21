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
