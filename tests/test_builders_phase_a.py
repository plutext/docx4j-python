"""CR-003 Phase A: the tree-layer builders.

Every builder is serialised with ``to_xml`` and compared against the XML it is
supposed to write, canonically --- ``scripts/canon.py``'s prefix-insensitive,
attribute-sorted normal form, which is the verdict the corpus round trip uses.
Two of the comparisons are against XML this repository did not write:

* the ``wp:inline`` comes from ``samples/Images.docx``, whose second picture is
  exactly what docx4j's ``BinaryPartAbstractImage.createImageInline`` produces,
  so :func:`inline_picture` is checked against docx4j's own output rather than
  against a copy of itself;
* the content-control kinds are read back from ``samples/invoice2013.docx``,
  which has a checkbox, a date picker, a picture control, two w15 repeating
  sections and twenty bindings, three of them ``w15:dataBinding``.

    .venv-fork/bin/python -m pytest tests/test_builders_phase_a.py -q
"""

from __future__ import annotations

import re
import struct
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon  # noqa: E402

from docx4j_py import wml as W  # noqa: E402
from docx4j_py.child import deep_copy_as, is_any_element  # noqa: E402
from docx4j_py.namespaces import WML_NS, declarations  # noqa: E402
from docx4j_py.traversal import RUN_HOLDERS, element_name, find, run_items_of, walk_all  # noqa: E402
from docx4j_py.w14 import el as w14_el  # noqa: E402
from docx4j_py.wml import el, to_xml, wml  # noqa: E402
from docx4j_py.wml.builders import (  # noqa: E402
    RPR_BASE_FIELDS,
    BuilderError,
    RPrElement,
    p,
    r,
    rpr_from_elements,
    rpr_to_elements,
    tbl,
    tc,
    tr,
)
from docx4j_py.wml.pictures import (  # noqa: E402
    DEFAULT_DPI,
    EMU_PER_INCH,
    emu_for,
    image_size,
    inline_picture,
)
from docx4j_py.wml.sdt import (  # noqa: E402
    ANY_NS,
    SDT_KINDS,
    W15_NS,
    W_NS,
    next_sdt_id,
    sdt,
    sdt_kind_of,
    sdt_pr,
    sdt_property,
)

SAMPLES = ROOT / "samples"
W_SDT = f"{{{WML_NS}}}sdt"


# ---------------------------------------------------------------------------
# comparing
# ---------------------------------------------------------------------------


def declared(fragment: str) -> str:
    """The fragment with every docx4j prefix declared, so lxml can parse it.

    ``to_xml``'s own output already declares what it uses, so a fragment that
    declares anything is left alone (declaring a prefix twice is not XML).
    """
    head, _, rest = fragment.partition(">")
    if "xmlns:" in head:
        return fragment
    closing = "/>" if head.endswith("/") else ">"
    return f"{head.removesuffix('/')} {declarations()}{closing}{rest}"


def same(built: str, expected: str) -> None:
    """Assert two fragments are canonically the same document."""
    report = canon.DiffReport()
    canon.diff_trees(
        canon.normalize_tree(declared(expected).encode()),
        canon.normalize_tree(built.encode()),
        report,
    )
    assert report.differences == {}, (built, expected, report.diffs[:5])


# ---------------------------------------------------------------------------
# tc and tr
# ---------------------------------------------------------------------------


def test_tc_of_a_string_is_one_paragraph():
    same(to_xml(tc("hello")), "<w:tc><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:tc>")


def test_tc_width_and_span():
    same(
        to_xml(tc("x", width=1234, span=2)),
        '<w:tc><w:tcPr><w:tcW w:w="1234" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>'
        "<w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>",
    )


def test_tc_with_no_width_has_no_tcPr():
    assert tc("x").tc_pr is None


def test_an_empty_cell_still_has_the_paragraph_word_requires():
    same(to_xml(tc([])), "<w:tc><w:p/></w:tc>")
    assert isinstance(tc([]).content[0], W.P)


def test_tc_takes_blocks_as_they_are():
    given = p("one")
    cell = tc([given, tbl([["a"]])])
    assert cell.content[0] is given
    assert given.parent is cell
    assert isinstance(cell.content[1], W.Tbl)


def test_tr_of_strings_with_widths():
    same(
        to_xml(tr(["a", "b"], widths=[100, 200])),
        "<w:tr>"
        '<w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="200" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>'
        "</w:tr>",
    )


def test_tr_header_repeats_the_row():
    same(
        to_xml(tr(["a"], header=True)),
        "<w:tr><w:trPr><w:tblHeader/></w:trPr>"
        "<w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr>",
    )


def test_tr_takes_a_built_cell_as_it_is():
    cell = tc("given", width=99)
    row = tr([cell, "made"], widths=[1, 2])
    assert row.content[0] is cell
    assert cell.tc_pr.tc_w.w == 99  # the widths argument does not overwrite it
    assert row.content[1].tc_pr.tc_w.w == 2


def test_tr_without_widths_writes_no_tcW():
    assert tr(["a", "b"]).content[0].tc_pr is None


def test_tbl_is_what_it_was_before_the_refactor():
    """``tbl`` over ``tr``/``tc`` writes what it wrote when it built cells itself."""
    same(
        to_xml(tbl([["a", "b"], ["c"]], style="TableGrid", widths=[1000, 2000])),
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="3000" w:type="dxa"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>'
        "<w:tr>"
        '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>'
        "</w:tr>"
        "<w:tr>"
        '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r/></w:p></w:tc>'
        "</w:tr></w:tbl>",
    )


def test_tr_and_tc_link_their_parents():
    row = tr(["a"])
    assert row.content[0].parent is row
    assert row.content[0].content[0].parent is row.content[0]


# ---------------------------------------------------------------------------
# inline_picture, against what docx4j writes
# ---------------------------------------------------------------------------


def docx_drawings(name: str) -> list[str]:
    """Every ``w:drawing`` in a sample's ``document.xml``, as source text."""
    with zipfile.ZipFile(SAMPLES / name) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
    return re.findall(r"<w:drawing>.*?</w:drawing>", xml, re.S)


#: The second picture of ``samples/Images.docx``: a ``wp:inline`` in exactly the
#: shape ``BinaryPartAbstractImage.createImageInline`` writes, which is what
#: :func:`inline_picture` is a port of. (The first one is Word's own, with the
#: extra ``a:srcRect``, ``a:picLocks``, ``bwMode`` and ``a:ln`` Word adds.)
DOCX4J_SHAPED_DRAWING = 1


#: ``BinaryPartAbstractImage.createImageInline``'s template, verbatim from
#: docx4j (the ``${...}`` are its own placeholders). This is the promise
#: :func:`inline_picture` makes, so it is compared character for character
#: rather than described again in the test.
DOCX4J_TEMPLATE = (
    '<wp:inline distT="0" distB="0" distL="0" distR="0">'
    '<wp:extent cx="${cx}" cy="${cy}"/>'
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:docPr id="${id1}" name="${filenameHint}" descr="${altText}"/>'
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
    "<a:graphic>"
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    "<pic:pic><pic:nvPicPr>"
    '<pic:cNvPr id="${id2}" name="${filenameHint}"/><pic:cNvPicPr/></pic:nvPicPr>'
    "<pic:blipFill>"
    '<a:blip r:embed="${rEmbedId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
    "</a:graphicData></a:graphic></wp:inline>"
)


def test_inline_picture_is_docx4js_template():
    """Element for element and attribute for attribute, docx4j's own string."""
    expected = DOCX4J_TEMPLATE
    for key, value in {
        "cx": "3238500",
        "cy": "2362200",
        "id1": "2",
        "id2": "0",
        "filenameHint": "pangolin.jpeg",
        "altText": "a pangolin",
        "rEmbedId": "rId6",
    }.items():
        expected = expected.replace("${" + key + "}", value)
    built = to_xml(
        inline_picture(
            "rId6",
            cx=3238500,
            cy=2362200,
            id=2,
            name="pangolin.jpeg",
            descr="a pangolin",
        )
    )
    same(built, f"<w:drawing>{expected}</w:drawing>")


def test_inline_picture_is_the_shape_in_a_real_document():
    """The same, against ``samples/Images.docx``'s docx4j-shaped picture.

    One attribute is deliberately substituted before the comparison: Word
    wrote ``wp:effectExtent l="19050"`` where docx4j's template writes zero
    (the ``//l=\"19050\"`` comment beside that line in
    ``BinaryPartAbstractImage`` is the fossil of the value it used to write).
    Everything else --- every element, every other attribute, and the two
    ``name`` attributes Word fills differently --- matches.
    """
    reference = docx_drawings("Images.docx")[DOCX4J_SHAPED_DRAWING]
    reference = reference.replace('<wp:effectExtent l="19050"', '<wp:effectExtent l="0"')
    built = to_xml(
        inline_picture(
            "rId6",
            cx=3238500,
            cy=2362200,
            id=2,
            name="Picture 1",          # wp:docPr/@name, as Word names a picture
            descr="pangolin.jpeg",     # wp:docPr/@descr, docx4j's altText
            pic_name="pangolin.jpeg",  # pic:cNvPr/@name, docx4j's filenameHint
        )
    )
    same(built, reference)


def test_inline_picture_returns_a_drawing_a_run_takes():
    drawing = inline_picture("rId1", cx=100, cy=200)
    assert isinstance(drawing, W.Drawing)
    run = r("before ", drawing)
    assert run.content[1] is drawing
    assert drawing.parent is run
    assert "<w:drawing>" in to_xml(run)


def test_inline_picture_defaults_name_and_descr_as_docx4j_does():
    drawing = inline_picture("rId1", cx=1, cy=2, id=7)
    inline = drawing.anchor_or_inline[0]
    assert inline.doc_pr.name == "Picture 7"
    assert inline.doc_pr.descr == ""
    assert inline.doc_pr.title is None
    # docx4j's id2, which Word writes as 0
    picture = inline.graphic.graphic_data.any_element[0]
    assert picture.nv_pic_pr.c_nv_pr.id == 0
    assert picture.nv_pic_pr.c_nv_pr.descr is None


def test_inline_picture_title_is_the_alt_text_title():
    inline = inline_picture("rId1", cx=1, cy=2, title="A pangolin").anchor_or_inline[0]
    assert inline.doc_pr.title == "A pangolin"
    assert 'title="A pangolin"' in to_xml(inline_picture("rId1", cx=1, cy=2, title="A pangolin"))


def test_inline_picture_can_link_rather_than_embed():
    drawing = inline_picture("rId1", cx=1, cy=2, link=True)
    blip = drawing.anchor_or_inline[0].graphic.graphic_data.any_element[0].blip_fill.blip
    assert (blip.link, blip.embed) == ("rId1", None)


def test_inline_picture_refuses_an_empty_extent():
    with pytest.raises(BuilderError) as raised:
        inline_picture("rId1", cx=0, cy=100)
    assert raised.value.code == "picture.empty_extent"
    assert "emu_for" in raised.value.hint


def test_inline_picture_parents_are_linked():
    drawing = inline_picture("rId1", cx=1, cy=2)
    inline = drawing.anchor_or_inline[0]
    assert inline.parent is drawing
    assert inline.graphic.parent is inline
    assert inline.extent.parent is inline


# ---------------------------------------------------------------------------
# image_size and emu_for
# ---------------------------------------------------------------------------


def media(name: str, suffix: str) -> bytes:
    """An image part of a sample docx, by extension."""
    with zipfile.ZipFile(SAMPLES / name) as archive:
        entry = next(
            n for n in archive.namelist() if "/media/" in n and n.lower().endswith(suffix)
        )
        return archive.read(entry)


def test_image_size_of_the_sample_png():
    info = image_size(media("Images.docx", ".png"))
    assert (info.width_px, info.height_px, info.format) == (340, 248, "png")
    assert (info.dpi_x, info.dpi_y) == (DEFAULT_DPI, DEFAULT_DPI)
    assert info.has_density is False


def test_image_size_of_the_sample_jpeg():
    info = image_size(media("Images.docx", ".jpeg"))
    assert (info.width_px, info.height_px, info.format) == (340, 248, "jpeg")


def test_a_png_with_a_phys_chunk_reports_its_density():
    info = image_size(media("2016_image_with_text_effects.docx", ".png"))
    assert (info.width_px, info.height_px) == (223, 146)
    assert info.has_density is True
    assert round(info.dpi_x) == 96


def test_emu_for_gives_word_its_own_extent_back():
    """The sample's ``wp:extent`` is what ``emu_for`` computes from the bytes."""
    info = image_size(media("Images.docx", ".png"))
    size = emu_for(info)
    assert (size.cx, size.cy) == (3238500, 2362200)
    assert size.scaled is False
    reference = docx_drawings("Images.docx")[DOCX4J_SHAPED_DRAWING]
    assert f'<wp:extent cx="{size.cx}" cy="{size.cy}"/>' in reference


def test_emu_for_scales_to_the_maximum_width_keeping_the_ratio():
    info = image_size(media("Images.docx", ".png"))
    natural = emu_for(info)
    size = emu_for(info, max_width_emu=1000000)
    assert size.scaled is True
    assert size.cx == 1000000
    assert size.cy == round(natural.cy * 1000000 / natural.cx)


def test_emu_for_leaves_an_image_that_fits_alone():
    info = image_size(media("Images.docx", ".png"))
    assert emu_for(info, max_width_emu=5486400) == emu_for(info)


def test_emu_is_pixels_over_dpi_times_914400():
    from docx4j_py.wml.pictures import ImageInfo

    size = emu_for(ImageInfo(96, 48, 96.0, 96.0, "png"))
    assert (size.cx, size.cy) == (EMU_PER_INCH, EMU_PER_INCH // 2)


#: A 3 by 5 GIF: the signature and the logical screen descriptor, which is all
#: :func:`image_size` reads. GIF carries no density at all.
TINY_GIF = b"GIF89a" + struct.pack("<HHBBB", 3, 5, 0, 0, 0)


def bmp(width: int, height: int, *, ppm: int = 0, core: bool = False) -> bytes:
    """A BMP with nothing but its headers, in either DIB flavour."""
    if core:
        header = struct.pack("<IHHHH", 12, width, height, 1, 24)
    else:
        header = struct.pack("<IiiHHIIiiII", 40, width, height, 1, 24, 0, 0, ppm, ppm, 0, 0)
    return b"BM" + struct.pack("<IHHI", 14 + len(header), 0, 0, 14 + len(header)) + header


def test_image_size_of_a_gif():
    info = image_size(TINY_GIF)
    assert (info.width_px, info.height_px, info.format) == (3, 5, "gif")
    assert (info.dpi_x, info.has_density) == (DEFAULT_DPI, False)


def test_image_size_of_a_bmp_with_a_density():
    info = image_size(bmp(7, 11, ppm=3780))
    assert (info.width_px, info.height_px, info.format) == (7, 11, "bmp")
    assert info.has_density is True
    assert round(info.dpi_x) == 96


def test_image_size_of_a_top_down_bmp_and_of_the_old_core_header():
    assert image_size(bmp(7, -11)).height_px == 11
    assert image_size(bmp(7, -11)).has_density is False
    core = image_size(bmp(4, 9, core=True))
    assert (core.width_px, core.height_px, core.has_density) == (4, 9, False)


def test_image_size_refuses_what_it_cannot_read():
    with pytest.raises(BuilderError) as raised:
        image_size(b"II\xbc\x01\x08\x00\x00\x00 not an image")
    assert raised.value.code == "image.unsupported_format"


def test_image_size_says_a_jpeg_is_truncated():
    with pytest.raises(BuilderError) as raised:
        image_size(media("Images.docx", ".jpeg")[:20])
    assert raised.value.code == "image.truncated"


def test_image_info_and_emu_size_are_json_ready():
    assert image_size(TINY_GIF).to_dict()["format"] == "gif"
    assert emu_for(image_size(TINY_GIF)).to_dict()["scaled"] is False


# ---------------------------------------------------------------------------
# the sdt family
# ---------------------------------------------------------------------------


def test_sdt_pr_writes_alias_tag_id_then_the_kind():
    same(
        to_xml(sdt_pr(kind="PlainText", tag="customer", title="Customer", id=42)),
        '<w:sdtPr><w:alias w:val="Customer"/><w:tag w:val="customer"/>'
        '<w:id w:val="42"/><w:text/></w:sdtPr>',
    )


def test_sdt_pr_lock_and_placeholder():
    same(
        to_xml(sdt_pr(lock="sdtContentLocked", placeholder="DefaultPlaceholder_1081868574")),
        '<w:sdtPr><w:lock w:val="sdtContentLocked"/>'
        '<w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder>'
        "</w:sdtPr>",
    )


def test_rich_text_is_the_untyped_control_word_writes():
    assert to_xml(sdt_pr(kind="RichText", id=1)).count("<w:") == 2  # sdtPr and id
    assert sdt_kind_of(sdt_pr(kind="RichText")) == "RichText"
    assert sdt_kind_of(None) == "RichText"


def test_the_checkbox_is_word_2010s_with_its_two_glyphs():
    same(
        to_xml(sdt_pr(kind="CheckBox")),
        "<w:sdtPr><w14:checkbox>"
        '<w14:checked w14:val="0"/>'
        '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
        '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>'
        "</w14:checkbox></w:sdtPr>",
    )


def test_the_date_picker_carries_a_format_and_a_mapping():
    same(
        to_xml(sdt_pr(kind="DatePicker")),
        '<w:sdtPr><w:date><w:dateFormat w:val="d/MM/yyyy"/>'
        '<w:storeMappedDataAs w:val="dateTime"/></w:date></w:sdtPr>',
    )


def test_the_repeating_section_is_w15():
    same(to_xml(sdt_pr(kind="RepeatingSection")), "<w:sdtPr><w15:repeatingSection/></w:sdtPr>")


@pytest.mark.parametrize("kind", SDT_KINDS)
def test_every_kind_reads_back_as_itself(kind):
    assert sdt_kind_of(sdt_pr(kind=kind)) == kind


def test_an_unknown_kind_says_which_ones_there_are():
    with pytest.raises(BuilderError) as raised:
        sdt_pr(kind="Barcode")
    assert raised.value.code == "sdt.unknown_kind"
    assert "RichText" in raised.value.hint


def test_title_and_alias_are_the_same_element():
    assert to_xml(sdt_pr(alias="T", id=1)) == to_xml(sdt_pr(title="T", id=1))
    with pytest.raises(BuilderError) as raised:
        sdt_pr(title="one", alias="other")
    assert raised.value.code == "sdt.title_and_alias"


#: The four forms, the class each takes, and the wrapper ``wml`` resolves the
#: name ``w:sdt`` under for it (CR-001 section 14.4 point 2).
FORMS = [
    ("block", W.SdtBlock, W.SdtContentBlock, None, lambda: [p("a")]),
    ("run", W.SdtRun, W.CTSdtContentRun, "p", lambda: [r("a")]),
    ("row", W.CTSdtRow, W.CTSdtContentRow, "tbl", lambda: [tr(["a"])]),
    ("cell", W.CTSdtCell, W.CTSdtContentCell, "tr", lambda: [tc("a")]),
]


@pytest.mark.parametrize("form,cls,content_cls,wrapper,build", FORMS, ids=[f[0] for f in FORMS])
def test_the_form_follows_the_content(form, cls, content_cls, wrapper, build):
    control = sdt(build(), id=1)
    assert isinstance(control, cls)
    assert isinstance(control.sdt_content, content_cls)
    assert element_name(control) == W_SDT


@pytest.mark.parametrize("form,cls,content_cls,wrapper,build", FORMS, ids=[f[0] for f in FORMS])
def test_every_form_round_trips_through_wml(form, cls, content_cls, wrapper, build):
    built = to_xml(sdt(build(), id=1, tag="t"), name=W_SDT)
    parsed = wml(built, wrapper=wrapper)
    assert isinstance(parsed, cls)
    same(to_xml(parsed, name=W_SDT), built)
    assert sdt_property(parsed.sdt_pr, "tag").val == "t"


def test_the_form_can_be_overridden():
    assert isinstance(sdt([p("a")], form="run"), W.SdtRun)


def test_an_unknown_form_says_which_ones_there_are():
    with pytest.raises(BuilderError) as raised:
        sdt([p("a")], form="footnote")
    assert raised.value.code == "sdt.unknown_form"


def test_a_repeating_section_is_not_a_run_level_control():
    with pytest.raises(BuilderError) as raised:
        sdt([r("a")], kind="RepeatingSection")
    assert raised.value.code == "sdt.form_mismatch"
    assert isinstance(sdt([p("a")], kind="RepeatingSection"), W.SdtBlock)


def test_sdt_content_and_parents():
    given = p("a")
    control = sdt([given], id=1)
    assert control.sdt_content.content[0] is given
    assert given.parent is control.sdt_content
    assert control.sdt_pr.parent is control


def test_sdt_property_takes_the_control_as_well_as_its_properties():
    control = sdt([p("a")], tag="t", id=3)
    assert sdt_property(control, "tag").val == "t"
    assert sdt_property(control.sdt_pr, "id").val == 3
    assert sdt_property(control, "dataBinding") is None
    assert sdt_property(None, "tag") is None


# -- against invoice2013.docx ------------------------------------------------


@pytest.fixture(scope="module")
def invoice():
    from docx4j_py.openpackaging import WordprocessingMLPackage

    with WordprocessingMLPackage.load(SAMPLES / "invoice2013.docx") as pkg:
        yield pkg.main_document_part.contents


@pytest.fixture(scope="module")
def invoice_properties(invoice):
    return find(invoice, W.SdtPr, mce="all")


def test_invoice2013_has_the_kinds_word_wrote(invoice_properties):
    from collections import Counter

    assert len(invoice_properties) == 22
    assert Counter(sdt_kind_of(pr) for pr in invoice_properties) == {
        "PlainText": 14,
        "RepeatingSection": 2,
        "RepeatingSectionItem": 2,
        "Picture": 1,
        "DatePicker": 1,
        "CheckBox": 1,
        "RichText": 1,
    }


def test_three_of_the_twenty_bindings_are_w15(invoice_properties):
    """CR-003 section 4: bindings are read from ``w:`` **and** ``w15:``."""
    in_w = [pr for pr in invoice_properties if sdt_property(pr, "dataBinding", W_NS)]
    in_w15 = [pr for pr in invoice_properties if sdt_property(pr, "dataBinding", W15_NS)]
    by_default = [pr for pr in invoice_properties if sdt_property(pr, "dataBinding")]
    assert (len(in_w), len(in_w15), len(by_default)) == (17, 3, 20)
    assert len([pr for pr in invoice_properties if sdt_property(pr, "dataBinding", ANY_NS)]) == 20
    binding = sdt_property(in_w15[0], "dataBinding", W15_NS)
    assert binding.xpath and binding.store_item_id


def test_the_w15_bindings_are_the_repeating_sections(invoice_properties):
    kinds = {
        sdt_kind_of(pr)
        for pr in invoice_properties
        if sdt_property(pr, "dataBinding", W15_NS) is not None
    }
    assert kinds <= {"RepeatingSection", "RichText", "PlainText"}


def test_next_sdt_id_is_one_above_the_highest_and_is_deterministic(invoice, invoice_properties):
    used = {
        getattr(sdt_property(pr, "id", W_NS), "val", None) for pr in invoice_properties
    } - {None}
    assert next_sdt_id(invoice) == max(used) + 1
    assert next_sdt_id(invoice) == next_sdt_id(invoice)


def test_next_sdt_id_of_a_tree_with_no_controls_is_one():
    assert next_sdt_id(p("nothing here")) == 1
    assert next_sdt_id(sdt([p("a")], id=5)) == 6


# ---------------------------------------------------------------------------
# rPr as an element list
# ---------------------------------------------------------------------------


def test_the_eg_rpr_base_members_are_the_schemas_fifty_one():
    names = [name for name, _qname in RPR_BASE_FIELDS]
    assert len(names) == 51
    assert names[0] == "r_style"
    assert "r_pr_change" not in names  # w:rPrChange is not a member of the group
    w14 = [q for _n, q in RPR_BASE_FIELDS if "2010/wordml" in q]
    assert len(w14) == 12  # the w14 text effects, glow through cntxtAlts
    assert RPR_BASE_FIELDS[-1][1].endswith("}cntxtAlts")


def full_rpr() -> W.RPr:
    """A ``w:rPr`` touching a toggle, a value, an enum and a w14 effect."""
    return el.rPr(
        r_style=el.rStyle(val="Strong"),
        b=el.b(),
        i=el.i(),
        color=el.color(val="FF0000"),
        sz=el.sz(val=24),
        u=el.u(val="single"),
        vert_align=el.vertAlign(val="superscript"),
        glow=w14_el.glow(rad=100),
    )


def test_rpr_to_elements_keeps_the_schema_order():
    elements = rpr_to_elements(full_rpr())
    assert [e.name for e in elements] == [
        "r_style",
        "b",
        "i",
        "color",
        "sz",
        "u",
        "vert_align",
        "glow",
    ]
    assert all(isinstance(e, RPrElement) for e in elements)
    assert elements[0].qname == f"{{{WML_NS}}}rStyle"
    assert elements[-1].to_dict()["value"] == "CTGlow"


def test_rpr_to_elements_deep_copies():
    rpr = full_rpr()
    copied = rpr_to_elements(rpr)[0].value
    assert copied is not rpr.r_style
    assert copied == rpr.r_style


def test_the_change_form_round_trips_and_writes_the_same_xml():
    rpr = full_rpr()
    original = rpr_from_elements(rpr_to_elements(rpr), cls=W.CtRprChangeRPr)
    assert isinstance(original, W.CtRprChangeRPr)
    # the change form keeps every member as a list, because EG_RPrBase is
    # unbounded there; the serialised XML is the same either way
    assert original.b == [el.b()]
    name = f"{{{WML_NS}}}rPr"
    same(to_xml(original, name=name), to_xml(rpr))


def test_an_rpr_object_can_be_given_instead_of_its_elements():
    rpr = full_rpr()
    assert rpr_from_elements(rpr, cls=W.CtRprChangeRPr) == rpr_from_elements(
        rpr_to_elements(rpr), cls=W.CtRprChangeRPr
    )
    assert rpr_from_elements(None) == W.RPr()


def test_the_round_trip_comes_back_to_an_rpr():
    rpr = full_rpr()
    there = rpr_from_elements(rpr, cls=W.CtRprChangeRPr)
    assert rpr_from_elements(there) == rpr


def test_a_paragraph_marks_properties_lose_their_revision_marks():
    para_rpr = el.para_r_pr(ins=el.ins(id=1, author="a"), b=el.b(), sz=el.sz(val=20))
    assert [e.name for e in rpr_to_elements(para_rpr)] == ["b", "sz"]
    original = rpr_from_elements(para_rpr, cls=W.CTParaRPrOriginal)
    assert original.ins is None
    assert original.b == [el.b()]


def test_a_full_rpr_change_serialises_as_word_writes_it():
    change = el.rPrChange(
        id=7,
        author="Claude",
        r_pr=rpr_from_elements(full_rpr(), cls=W.CtRprChangeRPr),
    )
    built = to_xml(change, name=f"{{{WML_NS}}}rPrChange")
    assert "xsi:type" not in built
    assert built.count("<w:rPr>") == 1


def test_rpr_from_elements_names_what_it_cannot_place():
    with pytest.raises(BuilderError) as raised:
        rpr_from_elements([("nonesuch", el.b())])
    assert raised.value.code == "rpr.unknown_property"
    with pytest.raises(BuilderError) as raised:
        rpr_from_elements([el.b()])
    assert raised.value.code == "rpr.not_an_element"


def test_pairs_may_name_the_field_the_local_name_or_the_qualified_name():
    for name in ("b_cs", "bCs", f"{{{WML_NS}}}bCs"):
        assert rpr_from_elements([(name, el.bCs())]).b_cs is not None


# ---------------------------------------------------------------------------
# deep_copy_as
# ---------------------------------------------------------------------------


def a_ppr() -> W.PPr:
    """A ``w:pPr`` with two ``PPrBase`` members and one only ``PPr`` has."""
    return el.pPr(
        p_style=el.pStyle(val="Heading1"),
        jc=el.jc(val="center"),
        r_pr=el.para_r_pr(b=el.b()),
    )


def test_a_ppr_in_a_ppr_change_would_otherwise_carry_xsi_type():
    change = el.pPrChange(id=1, author="Claude", p_pr=a_ppr())
    assert 'xsi:type="w:CT_PPr"' in to_xml(change, name=f"{{{WML_NS}}}pPrChange")


def test_deep_copy_as_writes_the_element_word_writes():
    change = el.pPrChange(id=1, author="Claude", p_pr=deep_copy_as(a_ppr(), W.PPrBase))
    built = to_xml(change, name=f"{{{WML_NS}}}pPrChange")
    assert "xsi:type" not in built
    assert "XMLSchema-instance" not in built
    same(
        built,
        '<w:pPrChange w:id="1" w:author="Claude">'
        '<w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>'
        "</w:pPrChange>",
    )


def test_deep_copy_as_keeps_only_what_the_target_declares():
    base = deep_copy_as(a_ppr(), W.PPrBase)
    assert type(base) is W.PPrBase
    assert base.p_style.val == "Heading1"
    assert not hasattr(base, "r_pr")  # PPrBase does not declare it


def test_deep_copy_as_copies_rather_than_moves():
    ppr = a_ppr()
    base = deep_copy_as(ppr, W.PPrBase)
    base.p_style.val = "Heading2"
    assert ppr.p_style.val == "Heading1"
    assert ppr.r_pr is not None


def test_deep_copy_as_links_the_copys_parents_and_takes_a_parent():
    change = el.pPrChange(id=1, author="a")
    base = deep_copy_as(a_ppr(), W.PPrBase, change)
    assert base.parent is change
    assert base.p_style.parent is base


def test_deep_copy_as_refuses_an_unrelated_class():
    with pytest.raises(TypeError, match="neither a base nor a sibling"):
        deep_copy_as(a_ppr(), W.CTBookmark)
    with pytest.raises(TypeError, match="generated class"):
        deep_copy_as(a_ppr(), dict)


# ---------------------------------------------------------------------------
# walk_all
# ---------------------------------------------------------------------------


@pytest.fixture(scope="module")
def wps_document():
    from docx4j_py.openpackaging import WordprocessingMLPackage

    with WordprocessingMLPackage.load(SAMPLES / "DrawingML_GraphicData_wps.docx") as pkg:
        yield pkg.main_document_part.contents


def test_walk_all_reaches_the_wildcards_inside_an_alternate_content(wps_document):
    typed: list[str | None] = []
    wildcards: list[tuple] = []
    walk_all(
        wps_document,
        lambda node, parent, name: typed.append(name),
        lambda node, parent, name: wildcards.append((name, node, parent)),
    )
    names = {name for name, _node, _parent in wildcards}
    # the wps shape inside the mc:Choice, which walk never hands over
    assert any(name.endswith("}wsp") for name in names)
    assert any("wordprocessingShape" in name for name in names)
    assert all(is_any_element(node) for _n, node, _p in wildcards)
    assert len(wildcards) > len(typed)


def test_walk_all_sees_the_attributes_on_a_wildcard(wps_document):
    """The reason ``walk`` is not enough: ``r:embed`` and friends live here."""
    seen: dict[str, dict] = {}
    walk_all(
        wps_document,
        lambda node, parent, name: None,
        lambda node, parent, name: seen.setdefault(name, dict(node.attributes or {})),
    )
    anchor = next(v for k, v in seen.items() if k.endswith("}anchor"))
    assert anchor["distL"] == "114300"
    assert any(key.endswith("}anchorId") for key in anchor)


def test_walk_all_reports_the_parent_of_a_wildcard(wps_document):
    parents: list[tuple] = []
    walk_all(
        wps_document,
        lambda node, parent, name: None,
        lambda node, parent, name: parents.append((name, parent)),
    )
    for name, parent in parents:
        assert parent is not None, name


def test_walk_all_without_a_wildcard_visitor_still_descends(wps_document):
    reached: list[str | None] = []
    walk_all(wps_document, lambda node, parent, name: reached.append(name))
    assert reached  # the typed nodes, the wildcards descended into silently


def test_walk_all_prunes_on_false():
    para = p("a", r("b"))
    seen: list[str | None] = []

    def visitor(node, parent, name):
        seen.append(name)
        return name != f"{{{WML_NS}}}r"

    walk_all(para, visitor)
    assert f"{{{WML_NS}}}t" not in seen


def test_walk_all_enters_both_branches_of_an_alternate_content():
    from docx4j_py.traversal import walk

    fragment = wml(
        "<w:r><mc:AlternateContent>"
        '<mc:Choice Requires="wps"><w:drawing/></mc:Choice>'
        "<mc:Fallback><w:pict/></mc:Fallback>"
        "</mc:AlternateContent></w:r>"
    )
    resolved: list[str | None] = []
    walk(fragment, lambda node, parent, name: resolved.append(name))
    everything: list[str | None] = []
    walk_all(fragment, lambda node, parent, name: everything.append(name))
    assert len(everything) >= len(resolved)


# ---------------------------------------------------------------------------
# run_items_of
# ---------------------------------------------------------------------------

#: Every holder class the model has for the names in :data:`RUN_HOLDERS`, and
#: the name each is written under. CR-001 section 14 warned that the field name
#: differs per holder; in this model it does not (see the Phase A notes).
HOLDERS = [
    (W.P, "p"),
    (W.PHyperlink, "hyperlink"),
    (W.RunIns, "ins"),
    (W.RunDel, "del"),
    (W.MoveFrom2, "moveFrom"),
    (W.MoveTo2, "moveTo"),
    (W.CTSmartTagRun, "smartTag"),
    (W.CTCustomXmlRun, "customXml"),
    (W.CTSimpleField, "fldSimple"),
    (W.CtDir, "dir"),
    (W.CtBdo, "bdo"),
    (W.CTSdtContentRun, "sdtContent"),
    (W.SdtContentBlock, "sdtContent"),
]


@pytest.mark.parametrize("cls,local", HOLDERS, ids=[c.__name__ for c, _ in HOLDERS])
def test_run_items_of_every_holder_is_its_own_live_list(cls, local):
    holder = cls()
    assert element_name(holder) == f"{{{WML_NS}}}{local}"
    items = run_items_of(holder)
    assert items is holder.content
    run = r("x")
    items.append(run)
    assert holder.content[-1] is run


def test_every_run_holder_name_has_a_class_that_answers():
    assert {name for _cls, name in ((c, f"{{{WML_NS}}}{n}") for c, n in HOLDERS)} | {
        f"{{{WML_NS}}}sdt"
    } == RUN_HOLDERS


@pytest.mark.parametrize("form,cls,content_cls,wrapper,build", FORMS, ids=[f[0] for f in FORMS])
def test_run_items_of_a_control_reaches_through_to_its_content(
    form, cls, content_cls, wrapper, build
):
    control = sdt(build(), id=1)
    assert run_items_of(control) is control.sdt_content.content


def test_run_items_of_anything_else_is_none():
    assert run_items_of(r("x")) is None
    assert run_items_of(W.Tbl()) is None
    assert run_items_of(tc("x")) is None
    assert run_items_of(W.SdtBlock()) is None  # no w:sdtContent at all


def test_run_items_of_a_parsed_paragraph_is_the_document_order():
    para = wml("<w:p><w:r><w:t>a</w:t></w:r><w:hyperlink><w:r><w:t>b</w:t></w:r></w:hyperlink></w:p>")
    items = run_items_of(para)
    assert [type(item).__name__ for item in items] == ["R", "PHyperlink"]
    assert [type(item).__name__ for item in run_items_of(items[1])] == ["R"]


def test_run_items_of_keeps_a_deletion_which_text_of_skips():
    """It is structural: the original view of a revision needs ``w:del``."""
    para = wml(
        "<w:p><w:r><w:t>kept</w:t></w:r>"
        '<w:del w:id="1" w:author="a"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>'
    )
    items = run_items_of(para)
    assert len(items) == 2
    deletion = items[1]
    assert element_name(deletion) == f"{{{WML_NS}}}del"
    assert len(run_items_of(deletion)) == 1
    assert W.text_of(para) == "kept"
