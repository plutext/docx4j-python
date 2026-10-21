"""Markdown from the text model: docx4j-markdown's tests, and the per-block form. CR-011 Phase D (section 14).

Ported from docx4j at ``198d015a3``, each with docx4j's expected string:

* ``MarkdownTextModelTest``, its 19 tests: HYPERLINK fields and anchors as
  links, hidden text, endnotes ``[^e1]``, front matter, comments in
  ``markup``, a text box after its anchor, ``w:tblHeader`` rows and
  alignment, an SVG picture, a picture with no image, numbering overrides
  and styles, and REQ-703's interrupted lists;
* ``MarkdownExportTest``, its 11 (the escaping test compares docx4j's
  string, as core-ts does, since it reads the Markdown back through docx4j's
  importer);
* ``MarkdownExportExtensionsTest``, its 9 (the two picture tests build the
  picture with this package's importer, as docx4j's builds it with its own);
* ``MarkdownAlternateContentTest`` (built, loaded kept, loaded with
  ``mce_preprocess``), ``MarkdownChartTest`` (both), and the two of
  ``MarkdownMathExportTest``'s six that read no LaTeX through docx4j's
  importer (``unsupportedOmmlFlattensWithWarning``,
  ``trackedChangesAcceptedRevisionsView``).

Then this package's own: :func:`markdown_blocks` joining to :func:`markdown`,
the destination callable, the per-block slices, and REQ-204.
"""

from __future__ import annotations

import base64
from pathlib import Path

from conformance import image_destinations
from lxml import etree
from mini_docx import NS, MiniDocx

from docx4j_py import create_package, load
from docx4j_py.model.text import (
    MarkdownExporter,
    MarkdownExportOptions,
    TextOptions,
    extract,
    markdown,
    markdown_blocks,
    plain_text,
)
from docx4j_py.openpackaging.load import LoadOptions

ROOT = Path(__file__).resolve().parent.parent.parent
CHARTS = ROOT / "tests" / "fixtures" / "charts"
PHASE0 = ROOT / "tests" / "fixtures" / "text-extraction" / "phase0"

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
_STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
_STYLES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
_NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
_NUMBERING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"


def _styles(extra: str = "", *, mono_default: bool = False) -> bytes:
    """docx4j's default styles (what ``createPackage`` writes), with `extra` styles added."""
    root = etree.fromstring(create_package().style_definitions_part.xml)
    if mono_default:
        fonts = root.find(f"{{{_W}}}docDefaults/{{{_W}}}rPrDefault/{{{_W}}}rPr/{{{_W}}}rFonts")
        fonts.attrib.clear()
        fonts.set(f"{{{_W}}}ascii", "Courier New")
        fonts.set(f"{{{_W}}}hAnsi", "Courier New")
    for style in etree.fromstring(f"<x xmlns:w=\"{_W}\">{extra}</x>"):
        root.append(style)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def docx(body: str, *, styles: str = "", mono_default: bool = False, numbering: str | None = None) -> MiniDocx:
    """A document as docx4j's ``createPackage`` makes it, with this body (``NS`` declared)."""
    made = MiniDocx.body(body).part("styles.xml", _STYLES_CT, _STYLES_REL, "rIdStyles", _styles(styles, mono_default=mono_default))
    if numbering is not None:
        made.part(
            "numbering.xml", _NUMBERING_CT, _NUMBERING_REL, "rIdNumbering", f"{_DECLARATION}<w:numbering {NS}>{numbering}</w:numbering>"
        )
    return made


def export(pkg, **options) -> str:
    """docx4j ``new MarkdownExporter().export(pkg)``, the text options by name."""
    return MarkdownExporter(MarkdownExportOptions(text_options=TextOptions(**options))).export(pkg)


def p(text: str) -> str:
    return f"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>"


# ---------------------------------------------------------------------------
# MarkdownTextModelTest
# ---------------------------------------------------------------------------


def test_a_hyperlink_field_is_a_link():
    pkg = docx(
        '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.org/a" \\l "part2" \\o "tip" </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>the example</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "<w:r><w:t>.</w:t></w:r></w:p>"
    ).load()
    assert export(pkg) == "See [the example](https://example.org/a#part2).\n"
    assert export(pkg, field_codes=True) == (
        'See  HYPERLINK "https://example.org/a" \\\\l "part2" \\\\o "tip" [the example](https://example.org/a#part2).\n'
    )


def test_a_simple_fields_hyperlink_and_an_anchor_are_links():
    pkg = docx(
        '<w:p><w:fldSimple w:instr=" HYPERLINK \\l &quot;intro&quot; "><w:r><w:t>back</w:t></w:r></w:fldSimple>'
        '<w:r><w:t xml:space="preserve"> and </w:t></w:r>'
        '<w:hyperlink w:anchor="end"><w:r><w:t>on</w:t></w:r></w:hyperlink></w:p>'
    ).load()
    assert export(pkg) == "[back](#intro) and [on](#end)\n"


def test_hidden_text_is_left_out():
    pkg = docx(
        '<w:p><w:r><w:t xml:space="preserve">shown </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>'
    ).load()
    assert export(pkg) == "shown \n", "nothing trimmed (WD REQ-602)"
    assert export(pkg, hidden=True) == "shown hidden\n"


def test_endnotes_are_notes():
    made = docx('<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>')
    made.part(
        "endnotes.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
        "rIdEn",
        f"{_DECLARATION}<w:endnotes {NS}>"
        '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>'
        '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>'
        '<w:endnote w:id="1"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> The end note.</w:t></w:r></w:p></w:endnote>'
        "</w:endnotes>",
    )
    assert export(made.load()) == "Body[^e1]\n\n[^e1]: The end note.\n"


def test_front_matter_under_metadata():
    made = docx(p("Body")).package_part(
        "docProps/core.xml",
        "application/vnd.openxmlformats-package.core-properties+xml",
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "rIdCore",
        f"{_DECLARATION}<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\""
        ' xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>A "quoted" title</dc:title></cp:coreProperties>',
    )
    pkg = made.load()
    md = export(pkg, metadata=True)
    assert md.startswith('---\ncore:\n  title: "A \\"quoted\\" title"\n'), md
    assert md.endswith("---\n\nBody\n"), md
    assert export(pkg) == "Body\n"


def test_a_comment_in_markup():
    made = docx(
        '<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>Claim</w:t></w:r><w:commentRangeEnd w:id="3"/>'
        '<w:r><w:commentReference w:id="3"/></w:r></w:p>'
    )
    made.part(
        "comments.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        "rIdCm",
        f'{_DECLARATION}<w:comments {NS}><w:comment w:id="3" w:author="Ann"><w:p><w:r><w:t>Check this</w:t></w:r></w:p></w:comment></w:comments>',
    )
    pkg = made.load()
    assert export(pkg) == "Claim\n"
    assert export(pkg, view="markup") == "Claim{>>Check this<<}\n"


def test_a_text_box_once_after_its_anchor():
    box = '<w:txbxContent><w:p><w:r><w:t>In the box</w:t></w:r></w:p></w:txbxContent>'
    pkg = docx(
        "<w:p><w:r><w:t>Anchor</w:t></w:r><w:r><mc:AlternateContent>"
        '<mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="Box"/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f"<wps:wsp><wps:txbx>{box}</wps:txbx><wps:bodyPr/></wps:wsp>"
        "</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>"
        f"<mc:Fallback><w:pict><v:shape><v:textbox>{box}</v:textbox></v:shape></w:pict></mc:Fallback>"
        "</mc:AlternateContent></w:r></w:p>" + p("After")
    ).load()
    assert export(pkg) == "Anchor\n\nIn the box\n\nAfter\n"


def _tc(jc: str | None, text: str | None) -> str:
    return (
        "<w:tc><w:p>"
        + ("" if jc is None else f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>')
        + ("" if text is None else f"<w:r><w:t>{text}</w:t></w:r>")
        + "</w:p></w:tc>"
    )


def test_header_rows_and_alignment():
    header = "<w:trPr><w:tblHeader/></w:trPr>"
    pkg = docx(
        "<w:tbl>"
        f"<w:tr>{header}{_tc('right', 'A')}{_tc(None, 'B')}</w:tr>"
        f"<w:tr>{header}{_tc(None, 'a2')}{_tc('center', 'b2')}</w:tr>"
        f"<w:tr>{_tc(None, '1')}{_tc('center', '2')}</w:tr></w:tbl>"
    ).load()
    assert export(pkg) == "|A / a2|B / b2|\n|---:|---|\n|1|2|\n"

    pkg = docx(
        f"<w:tbl><w:tr>{_tc('right', None)}{_tc('center', 'H')}</w:tr><w:tr>{_tc(None, 'x')}{_tc(None, '1')}</w:tr></w:tbl>"
    ).load()
    assert export(pkg) == "||H|\n|---:|:---:|\n|x|1|\n"


_SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="red"/></svg>'
_IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def _picture(blip: str) -> str:
    return (
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="381000" cy="190500"/><wp:docPr id="2" name="Picture 2" descr="A red box"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
        '<pic:nvPicPr><pic:cNvPr id="1" name=""/><pic:cNvPicPr/></pic:nvPicPr>'
        f"<pic:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        '<pic:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        "</wp:inline></w:drawing></w:r></w:p>"
    )


def test_an_svg_picture_with_no_png_copy_is_written_with_its_svg():
    pkg = (
        docx(
            _picture(
                '<a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
                '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/>'
                "</a:ext></a:extLst></a:blip>"
            )
        )
        .part("media/image1.svg", "image/svg+xml", _IMAGE_REL, "rIdSvg", _SVG)
        .load()
    )
    parts = [r.image.part for b in extract(pkg).blocks for r in b.runs if r.kind == "IMAGE"]
    assert [part.content_type for part in parts] == ["image/svg+xml"]
    md = export(pkg)
    prefix = "![A red box](data:image/svg+xml;base64,"
    assert md.startswith(prefix), md
    svg = base64.b64decode(md[len(prefix) : md.index(")")]).decode("utf-8")
    assert "<rect" in svg and 'fill="red"' in svg

    written: dict[str, bytes] = {}

    def to_files(image):
        written[image.file_name] = image.data
        return image.file_name

    md = MarkdownExporter(MarkdownExportOptions(image_destination=to_files)).export(pkg)
    assert md == "![A red box](image1.svg)\n"
    assert written["image1.svg"]


def test_a_picture_with_no_image_keeps_its_alt_text():
    pkg = docx(_picture('<a:blip r:embed="rIdMissing"/>')).load()
    assert export(pkg) == "![A red box]()\n"


def test_numbering_overrides():
    pkg = docx(
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Fifth</w:t></w:r></w:p>'
        "<w:p/>"
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Bulleted</w:t></w:r></w:p>',
        numbering='<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0">'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl></w:lvlOverride></w:num>',
    ).load()
    assert export(pkg) == "5. Fifth\n\n- Bulleted\n"


def test_numbering_through_styles():
    pkg = docx(
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Linked</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Numbered"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Other"/></w:pPr><w:r><w:t>Unnumbered</w:t></w:r></w:p>',
        styles='<w:style w:type="numbering" w:styleId="MyList"><w:name w:val="My List"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Numbered"><w:name w:val="Numbered"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Other"><w:name w:val="Other"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>',
        numbering='<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="MyList"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:pStyle w:val="Numbered"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        '<w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="MyList"/></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>',
    ).load()
    # two lists, III. and III. (docx4j CR-014 P2): the second takes the other delimiter
    assert export(pkg) == "3. Linked\n\n3) Styled\n\nUnnumbered\n"
    blocks = extract(pkg, list_labels=True).blocks
    assert [b.list_label for b in blocks] == ["III.\t", "III.\t", None]
    assert blocks[2].numbering is None


_TWO_LISTS = "".join(
    f'<w:abstractNum w:abstractNumId="{a}">'
    + "".join(
        f'<w:lvl w:ilvl="{lvl}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%{lvl + 1}."/>'
        f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{720 + 720 * lvl}" w:hanging="360"/></w:pPr></w:lvl>'
        for lvl in (0, 1)
    )
    + "</w:abstractNum>"
    for a in (1, 2)
) + '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>'


def _item(num_id: int, ilvl: int, text: str) -> str:
    return (
        f'<w:p><w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>'
        f"<w:r><w:t>{text}</w:t></w:r></w:p>"
    )


def _para(left: int | None, text: str) -> str:
    return (
        "<w:p>"
        + ("" if left is None else f'<w:pPr><w:ind w:left="{left}"/></w:pPr>')
        + ("" if not text else f"<w:r><w:t>{text}</w:t></w:r>")
        + "</w:p>"
    )


def _lists(*body: str):
    return docx("".join(body), numbering=_TWO_LISTS).load()


def test_an_interrupted_list_goes_on_at_its_count():
    pkg = _lists(_item(1, 0, "Apples"), _para(None, "A note"), _item(1, 0, "Pears"), _item(1, 0, "Plums"))
    assert export(pkg) == "1. Apples\n\nA note\n\n2. Pears\n3. Plums\n"
    blocks = extract(pkg).blocks
    assert blocks[0].numbering.value == 1
    assert blocks[2].numbering.value == 2


def test_an_indented_paragraph_continues_the_item():
    pkg = _lists(_item(1, 0, "Apples"), _para(720, "A note"), _item(1, 0, "Pears"))
    assert export(pkg) == "1. Apples\n   \n   A note\n\n2. Pears\n"


def test_a_paragraph_continues_the_deepest_item_it_reaches():
    pkg = _lists(_item(1, 0, "A"), _item(1, 1, "A.1"), _para(720, "about A"), _item(1, 1, "A.2"), _item(1, 0, "B"))
    assert export(pkg) == "1. A\n   \n   1. A.1\n   \n   about A\n   \n   2. A.2\n\n2. B\n"


def test_an_empty_paragraph_leaves_the_list_alone():
    pkg = _lists(_item(1, 0, "Apples"), _para(None, ""), _item(1, 0, "Pears"))
    assert export(pkg) == "1. Apples\n2. Pears\n"


def test_adjacent_lists_are_kept_apart():
    pkg = _lists(_item(1, 0, "X"), _item(2, 0, "Y"), _item(1, 0, "X again"))
    assert export(pkg) == "1. X\n\n1) Y\n\n2. X again\n"


def test_a_nested_list_resumes_under_its_parents_count():
    pkg = _lists(_item(1, 0, "A"), _item(1, 0, "B"), _item(1, 1, "B.1"), _para(None, "A note"), _item(1, 1, "B.2"))
    assert export(pkg) == "1. A\n2. B\n   1. B.1\n\nA note\n\n2. 2. B.2\n"


def test_the_markdown_of_a_result_is_of_that_reading():
    pkg = docx('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' + p("Body")).load()
    result = extract(pkg, view="accepted")
    assert MarkdownExporter().export(result) == export(pkg) == "# Title\n\nBody\n"
    assert markdown(result) == "# Title\n\nBody\n"


# ---------------------------------------------------------------------------
# MarkdownExportTest
# ---------------------------------------------------------------------------


def test_a_heading_by_its_outline_level_not_its_style_name():
    pkg = docx(
        '<w:p><w:pPr><w:pStyle w:val="Overskrift2"/></w:pPr><w:r><w:t>Localized heading</w:t></w:r></w:p>',
        styles='<w:style w:type="paragraph" w:styleId="Overskrift2"><w:name w:val="overskrift 2"/>'
        '<w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>',
    ).load()
    assert export(pkg) == "## Localized heading\n"


def test_a_headings_bold_does_not_become_markers():
    pkg = docx('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>The Title</w:t></w:r></w:p>').load()
    assert export(pkg) == "# The Title\n"


def test_a_fields_cached_result():
    pkg = docx(
        '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r><w:t xml:space="preserve"> of many.</w:t></w:r></w:p>'
    ).load()
    assert export(pkg) == "Page 7 of many.\n"


def test_a_content_control_contributes_its_content():
    pkg = docx(
        '<w:sdt><w:sdtPr><w:tag w:val="x"/></w:sdtPr><w:sdtContent>' + p("Inside the control") + "</w:sdtContent></w:sdt>"
    ).load()
    assert export(pkg) == "Inside the control\n"


def test_a_mono_font_run_is_inline_code():
    pkg = docx(
        '<w:p><w:r><w:t xml:space="preserve">run </w:t></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t>mono()</w:t></w:r></w:p>'
    ).load()
    assert export(pkg) == "run `mono()`\n"


def test_empty_paragraphs_are_dropped():
    assert export(docx(p("first") + "<w:p/>" + p("second")).load()) == "first\n\nsecond\n"


def test_markdown_significant_text_is_escaped():
    # docx4j reads its Markdown back through its importer; this is the string docx4j writes
    assert export(docx(p("2 * 3 = 6 and #tag [x]")).load()) == "2 \\* 3 = 6 and #tag \\[x\\]\n"


def test_a_mono_paragraph_style_is_a_fenced_code_block():
    pkg = docx(
        '<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">&lt;xpaths xmlns="http://opendope.org/xpaths"&gt;</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">  &lt;xpath id="x1"&gt;</w:t><w:br/>'
        '<w:t xml:space="preserve">    *not emphasis*</w:t></w:r></w:p>' + p("after"),
        styles='<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/>'
        '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/></w:rPr></w:style>',
    ).load()
    assert export(pkg) == (
        '```\n<xpaths xmlns="http://opendope.org/xpaths">\n  <xpath id="x1">\n    *not emphasis*\n```\n\nafter\n'
    )


def test_a_mono_document_default_stays_prose():
    assert export(docx(p("a * b"), mono_default=True).load()) == "a \\* b\n"


def test_a_label_level_is_not_a_list():
    pkg = docx(
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>The label is not a number.</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="8"/></w:numPr></w:pPr><w:r><w:t>But this is a list.</w:t></w:r></w:p>',
        numbering='<w:abstractNum w:abstractNumId="7"><w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="tab"/>'
        '<w:lvlText w:val="NOTE"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1134" w:hanging="1134"/></w:pPr></w:lvl></w:abstractNum>'
        '<w:abstractNum w:abstractNumId="8"><w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'
        '<w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num><w:num w:numId="8"><w:abstractNumId w:val="8"/></w:num>',
    ).load()
    assert export(pkg) == "**NOTE** The label is not a number.\n\n1. But this is a list.\n"


def test_toc_entries_are_dropped():
    pkg = docx(
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:hyperlink w:anchor="_Toc1004"><w:r><w:t>1.1</w:t></w:r><w:r><w:tab/></w:r>'
        "<w:r><w:t>Overview</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:hyperlink></w:p>"
        '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p>',
        styles='<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style>',
    ).load()
    assert export(pkg) == "# Overview\n"
    assert export(pkg, toc=True) == "[1.1\tOverview\t3](#_Toc1004)\n\n# Overview\n"


# ---------------------------------------------------------------------------
# MarkdownExportExtensionsTest
# ---------------------------------------------------------------------------


def _table(rows: str):
    return docx(f"<w:tbl>{rows}</w:tbl>").load()


def test_a_grid_span_pads_the_columns():
    pkg = _table(
        '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>span</w:t></w:r></w:p></w:tc></w:tr>'
        "<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>"
    )
    assert export(pkg) == "|span||\n|---|---|\n|a|b|\n"


def test_a_vertical_merges_continuation_is_empty():
    pkg = _table(
        "<w:tr><w:tc><w:p><w:r><w:t>h1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>h2</w:t></w:r></w:p></w:tc></w:tr>"
        '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>top</w:t></w:r></w:p></w:tc>'
        "<w:tc><w:p><w:r><w:t>r1</w:t></w:r></w:p></w:tc></w:tr>"
        "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>r2</w:t></w:r></w:p></w:tc></w:tr>"
    )
    assert export(pkg) == "|h1|h2|\n|---|---|\n|top|r1|\n||r2|\n"


def test_a_cell_of_several_paragraphs_flattens_with_br():
    pkg = _table("<w:tr><w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc></w:tr>")
    assert "one<br>two" in export(pkg)


_TRACKED = (
    '<w:p><w:r><w:t xml:space="preserve">keep </w:t></w:r>'
    '<w:ins w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins>'
    '<w:del w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve"> removed</w:delText></w:r></w:del>'
    "</w:p>"
)


def test_tracked_changes_accepted_by_default():
    assert export(docx(_TRACKED).load()) == "keep added\n"


def test_tracked_changes_in_markup():
    assert export(docx(_TRACKED).load(), view="markup") == "keep {++added++}{-- removed--}\n"


def test_tracked_changes_in_the_original_view():
    assert export(docx(_TRACKED).load(), view="original") == "keep  removed\n"


def test_tracked_changes_with_the_markup_detail():
    assert export(docx(_TRACKED).load(), view="markup", markup_detail="author") == "keep {++a: added++}{--a:  removed--}\n"


#: 1x1 px PNG.
_PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="


def _tiny():
    """A picture of docx4j's 1x1 PNG, its alt text ``tiny`` (docx4j builds it with its importer)."""
    picture = _picture('<a:blip r:embed="rIdPng"/>').replace('descr="A red box"', 'descr="tiny"')
    return docx(picture).part("media/image1.png", "image/png", _IMAGE_REL, "rIdPng", base64.b64decode(_PNG_BASE64)).load()


def test_an_image_written_where_the_destination_says():
    written: dict[str, bytes] = {}

    def to_images(image):
        written[image.file_name] = image.data
        return f"images/{image.file_name}"

    md = MarkdownExporter(MarkdownExportOptions(image_destination=to_images)).export(_tiny())
    assert "![tiny](images/image1.png)" in md
    assert len(written["image1.png"]) == len(base64.b64decode(_PNG_BASE64))


def test_an_image_is_a_data_uri_by_default():
    md = export(_tiny())
    assert f"![tiny](data:image/png;base64,{_PNG_BASE64})" in md
    assert "hyperlink" not in md


# ---------------------------------------------------------------------------
# MarkdownAlternateContentTest, MarkdownChartTest, MarkdownMathExportTest
# ---------------------------------------------------------------------------


def test_markdown_follows_the_text_model_through_alternate_content():
    made = docx(
        p("before")
        + f'<mc:AlternateContent><mc:Choice Requires="w14">{p("CHOICEW14")}</mc:Choice>'
        f'<mc:Fallback>{p("FALLBACKTEXT")}</mc:Fallback></mc:AlternateContent>'
        + p("after")
    )
    md = export(load(made.save(), options=LoadOptions(mce_preprocess=True)))
    assert md == "before\n\nCHOICEW14\n\nafter\n"

    # loaded kept (CR-012 Phase A: the schema admits the element at block level), the Markdown follows the
    # text model through it to the branch Word reads, as docx4j's does; nothing is skipped
    pkg = made.load()
    assert export(pkg) == "before\n\nCHOICEW14\n\nafter\n"
    assert not pkg.main_document_part.skipped


def test_a_pie_chart():
    md = export(load(CHARTS / "pie-chart.docx"))
    assert "\\[chart: pie\\]\n\nSales\n\n||Sales|\n|---|---:|\n|1st Qtr|8.2|\n|2nd Qtr|3.2|\n|3rd Qtr|1.4|\n|4th Qtr|1.2|\n" in md, md


def test_a_chartex_chart_from_its_choice_and_no_picture():
    md = export(load(CHARTS / "waterfall.docx"))
    assert "\\[chart: waterfall\\]\n\nChart Title\n\n||Series1|\n|---|---:|\n|Category 1|100|" in md, md
    assert "|Category 4|-40|" in md
    assert "![Chart 1]" not in md


def test_an_equation_outside_the_subset_is_its_runs_text():
    pkg = docx(
        "<w:p><m:oMath><m:func><m:fName><m:r><m:t>cosech</m:t></m:r></m:fName>"
        "<m:e><m:r><m:t>x</m:t></m:r></m:e></m:func></m:oMath></w:p>"
    ).load()
    assert export(pkg) == "$cosechx$\n"


def test_an_equations_tracked_changes_are_read_in_the_accepted_view():
    """``MarkdownMathExportTest.trackedChangesAcceptedRevisionsView`` (docx4j CR-010, issue 348)."""
    from docx4j_py.model.math import OmmlToLatex
    from docx4j_py.model.revisions import trees as rtrees
    from docx4j_py.wml import wml

    paragraph = wml(
        f"<w:p {NS}><m:oMath><m:r><m:t>a+</m:t></m:r>"
        '<m:r><w:ins w:id="1" w:author="a" w:date="2026-09-01T12:43:00Z"><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr>'
        "<m:t>b</m:t></w:ins></m:r>"
        '<m:r><w:del w:id="2" w:author="a" w:date="2026-09-01T12:43:00Z"><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr>'
        "<m:t>c</m:t></w:del></m:r></m:oMath></w:p>"
    )
    equation = next(item for item in paragraph.content if rtrees.is_math(item))
    assert OmmlToLatex().convert(equation) == "a+b"


# ---------------------------------------------------------------------------
# this package's: the blocks, the options, REQ-204
# ---------------------------------------------------------------------------


def test_the_blocks_join_to_the_markdown():
    for name in ("extraction-probe.docx", "extraction-probe-structure.docx", "extraction-probe-external.docm"):
        pkg = load(PHASE0 / name)
        for view in ("accepted", "original", "markup"):
            result = extract(pkg, view=view, metadata=True)
            blocks = markdown_blocks(result, pkg)
            assert blocks.front_matter + "".join(b.markdown for b in blocks.blocks) == blocks.markdown
            assert blocks.markdown == markdown(result, pkg)
            assert [b.block for b in blocks.blocks] == result.blocks


def test_a_blocks_markdown_is_whole_lines_from_its_first():
    pkg = _lists(_item(1, 0, "A"), _item(1, 1, "A.1"), _para(720, "about A"), _item(1, 1, "A.2"), _item(1, 0, "B"))
    blocks = markdown_blocks(extract(pkg), pkg).blocks
    assert [b.markdown for b in blocks] == ["1. A\n   \n", "   1. A.1\n   \n", "   about A\n   \n", "   2. A.2\n\n", "2. B\n"]


def test_a_code_paragraph_merged_into_a_fence_writes_nothing_and_says_so():
    pkg = docx(
        '<w:p><w:pPr><w:pStyle w:val="SourceCode"/></w:pPr><w:r><w:t>x = 1</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="SourceCode"/></w:pPr><w:r><w:t>y = 2</w:t></w:r></w:p>' + "<w:p/>" + p("after")
    ).load()
    blocks = markdown_blocks(extract(pkg), pkg).blocks
    assert [(b.markdown, b.merged) for b in blocks] == [
        ("```\nx = 1\ny = 2\n```\n\n", False),
        ("", True),
        ("", False),
        ("after\n", False),
    ]


def test_the_view_and_the_options_are_the_results():
    pkg = docx(_TRACKED).load()
    result = extract(pkg, view="markup")
    assert markdown(result) == "keep {++added++}{-- removed--}\n"
    assert plain_text(result) == "keep {++added++}{-- removed--}\n"
    options = MarkdownExportOptions()
    options.view = "original"
    assert MarkdownExporter(options).export(pkg) == "keep  removed\n"
    assert options.text_options.view == "original"


def test_markdown_changes_nothing(tmp_path):
    """REQ-204: a Markdown export in every view, with every story, saves byte for byte as before."""
    for name in ("extraction-probe.docx", "extraction-probe-structure.docx", "extraction-probe-inspect.docx"):
        pkg = load(PHASE0 / name)
        before = pkg.save()
        for view in ("accepted", "original", "markup"):
            export(
                pkg, view=view, include="headers, footers, comments, glossary", hidden=True, field_codes=True,
                list_labels=True, metadata=True,
            )
        assert not [part for part in pkg.parts.values() if getattr(part, "is_unmarshalled", False)]
        assert pkg.save() == before, name


def test_the_markdown_goldens_are_exact():
    for name in ("extraction-probe", "extraction-probe-structure", "extraction-probe-external"):
        source = PHASE0 / (name + (".docm" if name.endswith("external") else ".docx"))
        golden = (PHASE0 / f"{name}.md").read_bytes().decode("utf-8")
        pkg = load(source)
        assert image_destinations(markdown(extract(pkg), pkg)) == image_destinations(golden), name


# ---------------------------------------------------------------------------
# found by the byte check against docx4j (CR-011 section 14)
# ---------------------------------------------------------------------------


def test_the_core_properties_dates_are_as_the_part_writes_them():
    """docx4j reads ``created`` and ``modified`` as their literal's text and ``lastPrinted`` by ``toXMLFormat``."""
    made = docx(p("Body")).package_part(
        "docProps/core.xml",
        "application/vnd.openxmlformats-package.core-properties+xml",
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "rIdCore",
        f"{_DECLARATION}<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\""
        ' xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dcterms:created xsi:type="dcterms:W3CDTF">2018-11-28T22:07:38.4950271Z</dcterms:created>'
        "<cp:lastPrinted>2012-01-13T01:57:00Z</cp:lastPrinted></cp:coreProperties>",
    )
    pkg = made.load()
    core = extract(pkg).metadata["core"]
    assert core == {"created": "2018-11-28T22:07:38.4950271Z", "lastPrinted": "2012-01-13T01:57:00Z"}
    assert export(pkg, metadata=True).startswith(
        '---\ncore:\n  created: "2018-11-28T22:07:38.4950271Z"\n  lastPrinted: "2012-01-13T01:57:00Z"\n---\n\n'
    )


def test_a_table_given_as_content_is_one_table_block():
    """docx4j's ``Texts.of(pkg, tbl)``: its blockItems gives a table none, so the table is the one block read."""
    pkg = create_package()
    [table] = pkg.body.insert_xml(f"<w:tbl><w:tblPr/><w:tblGrid/><w:tr>{_tc(None, 'A')}{_tc(None, 'B')}</w:tr></w:tbl>")
    blocks = extract(table.element, package=pkg).blocks
    assert [(b.kind, b.address) for b in blocks] == [("TABLE", "content/0")]
    assert markdown(extract(table.element)) == "|A|B|\n|---|---|\n"


def test_a_list_levels_text_and_start_are_the_definitions():
    """The block JSON's ``levelText`` and ``start``, which the level's ``lvl_text`` misread had left None and 1."""
    numbering = (
        '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="4"/><w:numFmt w:val="lowerRoman"/>'
        '<w:lvlText w:val="(%1)"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>'
    )
    pkg = docx(_item(1, 0, "Fourth"), numbering=numbering).load()
    [block] = extract(pkg).blocks
    assert block.numbering.to_dict() == {
        "numId": "1", "level": 0, "format": "lowerRoman", "levelText": "(%1)", "start": 4, "value": 4, "abstractNumId": "1",
    }
    assert export(pkg) == "4. Fourth\n"
