"""VML in the model: CR-012 Phase C.

    .venv-fork/bin/python -m pytest tests/openpackaging/test_vml.py -q

Until Phase C the schema copy had no VML, so below a ``w:pict`` or a
``w:object`` (whose content models are lax wildcards naming the VML and ``o:``
namespaces) and in ``settings.xml``'s ``w:shapeDefaults`` everything was an
``AnyElement`` tree. docx4j's five VML schemas, at the schema copy's commit,
are in the copy now (``schemas/vml/``), and the parser types VML wherever it
stands, in a branch or not. What is held (CR-012 section 4.3 and section 12):

* VML typed under ``w:pict`` and ``w:object``, in an ``mc:Fallback`` and
  outside any branch; ``v:textbox``'s ``w:txbxContent`` typed with it;
* ``settings.xml``'s ``o:shapedefaults`` and ``o:shapelayout`` typed;
* a header's VML watermark typed;
* each part round-trips canonically identical;
* ``w:object``'s and ``w:pict``'s children, one list (``any_and_any``,
  docx4j's ``getAnyAndAny()``), saved in document order (section 12.15);
* VML's booleans (``ST_TrueFalse``, a string enumeration) keep their
  spelling, and ``o:spt`` (``xsd:float``) and ``o:gfxdata``
  (``xsd:base64Binary``) theirs, through the fork's ``lexical_types``;
* ``insert_ooxml`` rewrites a typed ``v:imagedata``'s ``r:id``;
* the census on ``samples/``, pinned;
* the islands are still reached below a wildcard node, and nowhere else.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

import canon
import vml_census

from docx4j_py import load
from docx4j_py.child import is_any_element
from docx4j_py.mce import AlternateContent
from docx4j_py.traversal import find, walk_all
from docx4j_py.vml.main import (
    CTFill,
    CTImageData,
    CTRect,
    CTShape,
    CTShapetype,
    CTTextbox,
    CTTextPath,
)
from docx4j_py.vml.officedrawing import (
    CTIdMap,
    CTOLEObject,
    CTShapeDefaults,
    CTShapeLayout,
)
from docx4j_py.vml.wordprocessing_drawing import CTWrap
from docx4j_py.wml import CTObject, CTTxbxContent, Pict, R, wml
from docx4j_py.wml import CTShapeDefaults as ShapeDefaultsHolder

SAMPLES = ROOT / "samples"
FIXTURES = ROOT / "tests" / "fixtures"
INSPECT = FIXTURES / "text-extraction" / "phase0" / "extraction-probe-inspect.docx"
WPS = SAMPLES / "DrawingML_GraphicData_wps.docx"
TRACKED = FIXTURES / "revisions" / "cr043" / "cr043-tracked.docx"
WATERMARKED = FIXTURES / "loadAndSave.docx"


def wildcards(root) -> list[str]:
    """Every wildcard node below `root`, by name."""
    found: list[str] = []
    walk_all(root, lambda node, parent, name: None, lambda node, parent, name: found.append(name))
    return found


def outside_any_branch(node) -> bool:
    parent = node.parent
    while parent is not None:
        if isinstance(parent, AlternateContent):
            return False
        parent = parent.parent
    return True


# ---------------------------------------------------------------------------
# typed where it stands
# ---------------------------------------------------------------------------


def test_vml_under_a_pict_outside_any_branch_is_typed():
    """The inspection probe's header holds a legacy ``w:pict`` straight in a run: its VML is typed."""
    header = next(p for p in load(INSPECT).parts.values() if str(p.part_name) == "/word/header2.xml")
    (pict,) = find(header.contents, Pict)
    assert isinstance(pict.parent, R) and outside_any_branch(pict)
    shapetype, shape = pict.any_and_any
    assert isinstance(shapetype, CTShapetype) and isinstance(shape, CTShape)
    assert wildcards(header.contents) == []


def test_vml_under_a_pict_in_a_fallback_is_typed():
    """The wps shape's Fallback: a ``v:rect`` holding a ``v:textbox`` and its ``w:txbxContent``, all typed."""
    contents = load(WPS).main_document_part.contents
    (ac,) = find(contents, AlternateContent, mce="all")
    (pict,) = ac.fallback.any_element
    assert isinstance(pict, Pict)
    (rect,) = pict.any_and_any
    assert isinstance(rect, CTRect)
    assert (rect.spid, rect.anchor_id) == ("_x0000_s1026", "2953E3D1")
    assert wildcards(contents) == []


def test_vml_under_an_object_is_typed():
    """A ``w:object``'s ``v:shapetype``, ``v:shape`` and ``o:OLEObject``, typed, the object's ProgID a field."""
    objects = find(load(INSPECT).main_document_part.contents, CTObject)
    assert len(objects) == 2
    for obj in objects:
        assert outside_any_branch(obj)
        shapetype, shape, ole = obj.any_and_any
        assert [isinstance(c, k) for c, k in ((shapetype, CTShapetype), (shape, CTShape), (ole, CTOLEObject))] == [True] * 3
        assert ole.prog_id in ("Word.Document.12", "Excel.Sheet.12")
        assert ole.id and ole.id.startswith("rId")
        (image,) = [c for c in shape.content if isinstance(c, CTImageData)]
        assert image.schemas_openxmlformats_org_office_document_2006_relationships_id.startswith("rId")


def test_a_vml_text_box_holds_a_typed_txbx_content():
    """``v:textbox`` (``CT_Textbox``) declares its ``w:txbxContent``: a ``CTTxbxContent`` with paragraphs."""
    contents = load(WPS).main_document_part.contents
    (ac,) = find(contents, AlternateContent, mce="all")
    (textbox,) = find(ac.fallback, CTTextbox, mce="all")
    content = textbox.txbx_content_or_local_element
    assert isinstance(content, CTTxbxContent)
    assert content.content and all(not is_any_element(item) for item in content.content)


def test_the_tracked_vml_text_box_is_typed_down_to_its_changes():
    """cr043's C22: the Fallback's VML text box holds the tracked deletion and insertion, typed (CR-010's twins)."""
    from docx4j_py.wml import RunDel, RunIns

    contents = load(TRACKED).main_document_part.contents
    fallback_boxes = [
        box
        for ac in find(contents, AlternateContent, mce="all")
        if ac.fallback is not None
        for box in find(ac.fallback, CTTxbxContent, mce="all")
    ]
    assert fallback_boxes
    assert any(find(box, RunIns) and find(box, RunDel) for box in fallback_boxes)
    assert wildcards(contents) == []


def test_settings_shape_defaults_and_layout_are_typed():
    settings = load(SAMPLES / "Symbols.docx").main_document_part.document_settings_part.contents
    defaults = settings.shape_defaults
    assert isinstance(defaults, ShapeDefaultsHolder)  # w:shapeDefaults, whose o: wildcard holds the two
    shape_defaults, shape_layout = defaults.schemas_microsoft_com_office_office_element
    assert isinstance(shape_defaults, CTShapeDefaults)
    assert shape_defaults.spidmax == 1026
    assert isinstance(shape_layout, CTShapeLayout)
    assert isinstance(shape_layout.idmap, CTIdMap) and shape_layout.idmap.data == "1"
    assert wildcards(settings) == []


@pytest.mark.parametrize(("path", "string"), [(WATERMARKED, "SAMPLE"), (INSPECT, "WMK63_WATERMARK")], ids=["loadAndSave", "inspect"])
def test_a_header_watermark_is_typed(path, string):
    """Word's text watermark: a ``v:shape`` with its ``v:textpath`` string (and ``v:fill``, ``w10:wrap``), typed."""
    pkg = load(path)
    headers = [p for p in pkg.parts.values() if str(p.part_name).startswith("/word/header")]
    shapes = [
        shape
        for header in headers
        for shape in find(header.contents, CTShape)
        if (shape.id or "").startswith("PowerPlusWaterMarkObject")
    ]
    assert shapes
    for shape in shapes:
        (textpath,) = [item for item in shape.content if isinstance(item, CTTextPath)]
        assert textpath.string == string
    for header in headers:
        assert wildcards(header.contents) == []
    if path == INSPECT:
        (shape,) = shapes
        assert any(isinstance(item, CTFill) for item in shape.content)
        assert any(isinstance(item, CTWrap) for item in shape.content)


# ---------------------------------------------------------------------------
# the round trip
# ---------------------------------------------------------------------------

#: The inspection probe's main part, whose two ``w:object`` were saved in the
#: schema's order until section 12.15's patch, is canonically identical now.
ROUND_TRIP = [
    (INSPECT, "/word/header2.xml"),
    (INSPECT, "/word/document.xml"),
    (WPS, "/word/document.xml"),
    (TRACKED, "/word/document.xml"),
    (WATERMARKED, "/word/header1.xml"),
    (WATERMARKED, "/word/document.xml"),
    (SAMPLES / "Symbols.docx", "/word/settings.xml"),
    (FIXTURES / "numbering-stories.docx", "/word/document.xml"),
]


@pytest.mark.parametrize(("path", "part_name"), ROUND_TRIP, ids=[f"{p.name}{n}" for p, n in ROUND_TRIP])
def test_each_round_trips_canonically_identical(path, part_name):
    pkg = load(path)
    part = next(p for p in pkg.parts.values() if str(p.part_name) == part_name)
    source = part.xml
    saved = part._marshal(part.contents)
    report = canon.compare(source, saved)
    assert report.identical, report.differences
    assert part.skipped == []
    assert b"xsi:type" not in saved


# ---------------------------------------------------------------------------
# document order (CR-012 section 12.15)
# ---------------------------------------------------------------------------

PROBE = FIXTURES / "text-extraction" / "phase0" / "extraction-probe.docx"

#: Prefixes for the names the order tests compare.
_PREFIX = {
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main": "w",
    "urn:schemas-microsoft-com:vml": "v",
    "urn:schemas-microsoft-com:office:office": "o",
}


def children_of(xml: bytes, holder: str) -> list[list[str]]:
    """Each ``holder`` (``w:object``, ``w:pict``) in ``xml``: its children, as ``prefix:local``."""
    from lxml import etree

    ns, _, local = holder.partition(":")
    uri = next(u for u, p in _PREFIX.items() if p == ns)
    out = []
    for element in etree.fromstring(xml).iter(f"{{{uri}}}{local}"):
        names = []
        for child in element:
            qname = etree.QName(child)
            names.append(f"{_PREFIX.get(qname.namespace, qname.namespace)}:{qname.localname}")
        out.append(names)
    return out


def saved_part(package_bytes: bytes, name: str = "word/document.xml") -> bytes:
    import io
    import zipfile

    with zipfile.ZipFile(io.BytesIO(package_bytes)) as zf:
        return zf.read(name)


@pytest.mark.parametrize("path", [INSPECT, PROBE], ids=["inspect", "probe"])
def test_an_objects_children_are_saved_in_document_order(path):
    """``w:object``'s ``v:shapetype``, ``v:shape``, ``o:OLEObject`` survive a load, unmarshal and save in that order.

    ``CT_PictureBase`` was a repeating sequence of a VML wildcard and an ``o:``
    wildcard, two lists the serialiser wrote interleaved: ``o:OLEObject``
    before the ``v:shape`` its ``ShapeID`` names, which Word shows and will
    not activate (section 12.14, file 3). Patch 6 makes it one wildcard list
    in document order. (The probe's fifth object is a control: ``v:shapetype``,
    ``v:shape``, ``w:control``.)
    """
    pkg = load(path)
    part = pkg.main_document_part
    source = part.xml
    in_source = children_of(source, "w:object")
    three = [c for c in in_source if c == ["v:shapetype", "v:shape", "o:OLEObject"]]
    assert three
    part.contents  # unmarshalled, so the save re-marshals it
    package_bytes = pkg.save()
    saved = saved_part(package_bytes)
    assert children_of(saved, "w:object") == in_source
    report = canon.compare(source, saved)
    assert report.identical, report.differences
    assert b"xsi:type" not in saved
    objects = find(load(package_bytes).main_document_part.contents, CTObject)
    typed = [[type(c) for c in obj.any_and_any] for obj in objects if len(obj.any_and_any) == 3]
    assert len(typed) == len(three)
    for classes in typed:
        assert [issubclass(c, k) for c, k in zip(classes, (CTShapetype, CTShape, CTOLEObject))] == [True] * 3


#: A ``w:pict`` whose VML and ``o:`` children are not interleaved, in two orders:
#: every ``v:`` then every ``o:``, and the other way about. Written interleaved, the
#: first came back ``v:shapetype``, ``o:signatureline``, ``v:shape``, ``o:lock``, ``v:rect``.
PICT_ORDERS = {
    "v-then-o": (
        '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75"/>'
        '<v:shape id="s1" type="#_x0000_t75" style="width:10pt;height:10pt"/>'
        '<v:rect id="s2" style="width:10pt;height:10pt"/>'
        '<o:signatureline v:ext="edit" id="{00000000-0000-0000-0000-000000000001}" issignatureline="t"/>'
        '<o:lock v:ext="edit" aspectratio="t"/>',
        ["v:shapetype", "v:shape", "v:rect", "o:signatureline", "o:lock"],
    ),
    "o-then-v": (
        '<o:lock v:ext="edit" aspectratio="t"/>'
        '<o:signatureline v:ext="edit" id="{00000000-0000-0000-0000-000000000002}" issignatureline="t"/>'
        '<v:shape id="s3" style="width:10pt;height:10pt"/>'
        '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"/>',
        ["o:lock", "o:signatureline", "v:shape", "v:shapetype"],
    ),
}


@pytest.mark.parametrize("order", PICT_ORDERS, ids=list(PICT_ORDERS))
def test_a_picts_children_are_saved_in_document_order(order):
    """A ``w:pict`` mixing ``v:`` and ``o:`` children keeps their order through a save, a load, an unmarshal and a save."""
    from docx4j_py import create_package

    inner, expected = PICT_ORDERS[order]
    pkg = create_package()
    pkg.main_document_part.contents.body.content.insert(0, wml(f"<w:p><w:r><w:pict>{inner}</w:pict></w:r></w:p>"))
    package_bytes = pkg.save()
    first = saved_part(package_bytes)
    assert children_of(first, "w:pict") == [expected]

    again = load(package_bytes)
    part = again.main_document_part
    (pict,) = find(part.contents, Pict)
    assert [_PREFIX[c.Meta.namespace] for c in map(type, pict.any_and_any)] == [n.split(":")[0] for n in expected]
    assert wildcards(part.contents) == []
    second = saved_part(again.save())
    assert children_of(second, "w:pict") == [expected]
    assert canon.compare(first, second).identical
    assert part.skipped == []


# ---------------------------------------------------------------------------
# spellings kept
# ---------------------------------------------------------------------------

BOOLEANS = (
    '<w:r><w:pict><v:shape id="s" filled="t" stroked="f" o:allowincell="true" o:allowoverlap="false"'
    ' o:preferrelative="t"><v:textpath on="true" fitshape="f" string="x"/></v:shape></w:pict></w:r>'
)


def test_the_boolean_spellings_are_kept():
    """``ST_TrueFalse`` is a string enumeration (``t``, ``f``, ``true``, ``false``): each is written as read.

    The model writes an ``xsd:boolean`` numerically (CR-001 Q5); VML declares
    none, so nothing of VML's is respelled.
    """
    run = wml(BOOLEANS)
    (shape,) = find(run, CTShape)
    assert [shape.filled.value, shape.stroked.value, shape.allowincell.value, shape.allowoverlap.value] == [
        "t", "f", "true", "false",
    ]
    out = canon.normalize(run_xml(run))
    for name, value in (("filled", "t"), ("stroked", "f"), ("allowincell", "true"), ("allowoverlap", "false"), ("on", "true"), ("fitshape", "f")):
        assert f'{name}="{value}"'.encode() in out, name


def run_xml(run) -> bytes:
    from docx4j_py.fragments import to_xml

    return to_xml(run).encode("utf-8")


def test_o_spt_and_o_gfxdata_keep_their_spelling():
    """``o:spt="202"`` is not saved ``202.0``, and Word's wrapped ``o:gfxdata`` is not saved on one line.

    The fork's ``ParserConfig(lexical_types=)`` (stage 12, item 28): a float or
    bytes value the converter would spell otherwise keeps the source's. A value
    set in code takes the converter's spelling.
    """
    from docx4j_xsdata.models.datatype import XmlLexicalBytes, XmlLexicalFloat

    header = next(p for p in load(INSPECT).parts.values() if str(p.part_name) == "/word/header2.xml")
    source = header.xml
    (shapetype,) = find(header.contents, CTShapetype)
    assert type(shapetype.spt) is XmlLexicalFloat and shapetype.spt == 136.0
    saved = header._marshal(header.contents)
    assert b'o:spt="136"' in saved and b'o:spt="136.0"' not in saved
    assert canon.compare(source, saved).identical
    shapetype.spt = 75.0
    assert b'o:spt="75.0"' in header._marshal(header.contents)

    part = load(WPS).main_document_part
    source = part.xml
    (rect,) = find(part.contents, CTRect, mce="all")
    assert type(rect.gfxdata) is XmlLexicalBytes and "\n" in rect.gfxdata.lexical
    assert rect.gfxdata.startswith(b"PK")  # the shape's DrawingML, zipped
    assert canon.compare(source, part._marshal(part.contents)).identical


# ---------------------------------------------------------------------------
# a reader of references
# ---------------------------------------------------------------------------


def test_insert_ooxml_rewrites_a_typed_imagedatas_relationship_id():
    """``v:imagedata/@r:id`` is a typed field now; the content API's rewrite reads relationships-namespace fields."""
    from docx4j_py.model.content.ooxml import rewrite_relationship_ids

    run = wml(
        '<w:r><w:pict><v:shape id="p"><v:imagedata r:id="rId7" o:title=""/></v:shape></w:pict></w:r>'
    )
    rewrite_relationship_ids(run, {"rId7": "rId12"})
    (image,) = find(run, CTImageData)
    assert image.schemas_openxmlformats_org_office_document_2006_relationships_id == "rId12"
    assert image.id is None  # VML's own id attribute, which is no relationship


# ---------------------------------------------------------------------------
# the census on samples/, pinned
# ---------------------------------------------------------------------------


def test_the_census_on_samples():
    """Every VML element in ``samples/`` is typed: in ``settings.xml`` and in the three Fallbacks."""
    total: dict = {}
    for path in sorted(SAMPLES.glob("*.docx")):
        pkg = load(path)
        for _name, part in vml_census.wml_parts(pkg):
            for (label, kind, where), count in vml_census.census(part.contents)["elements"].items():
                key = (type(part).__name__, label, kind, where)
                total[key] = total.get(key, 0) + count
    assert total == {
        ("DocumentSettingsPart", "o:idmap", "typed", "outside"): 12,
        ("DocumentSettingsPart", "o:shapedefaults", "typed", "outside"): 13,
        ("DocumentSettingsPart", "o:shapelayout", "typed", "outside"): 12,
        ("HeaderPart", "v:oval", "typed", "fallback"): 1,
        ("HeaderPart", "w:pict", "typed", "fallback"): 1,
        ("MainDocumentPart", "v:oval", "typed", "fallback"): 1,
        ("MainDocumentPart", "v:rect", "typed", "fallback"): 1,
        ("MainDocumentPart", "v:textbox", "typed", "fallback"): 1,
        ("MainDocumentPart", "w:pict", "typed", "fallback"): 2,
        ("MainDocumentPart", "w:txbxContent", "typed", "choice"): 1,
        ("MainDocumentPart", "w:txbxContent", "typed", "fallback"): 1,
    }


# ---------------------------------------------------------------------------
# the islands: still reached below a wildcard node
# ---------------------------------------------------------------------------

#: A text box in a VML-namespace element no class binds: everything below a
#: wildcard node is a wildcard, so its ``w:txbxContent`` is one too.
BELOW_A_WILDCARD = (
    "<w:p><w:r><w:t>before</w:t></w:r><w:r><w:pict><v:unbound><v:textbox><w:txbxContent><w:p>"
    '<w:ins w:id="10" w:author="A" w:date="2026-10-22T00:00:00Z"><w:r><w:t>boxed</w:t></w:r></w:ins>'
    "</w:p></w:txbxContent></v:textbox></v:unbound></w:pict></w:r></w:p>"
)

#: A branch's ``w:moveTo``, neither global nor a docx4j root: a wildcard tree,
#: with the ``w:pict`` and its VML below it.
IN_A_BRANCHS_MOVE = (
    '<w:p><w:r><w:t>before</w:t></w:r><mc:AlternateContent><mc:Choice Requires="w14">'
    '<w:moveTo w:id="9" w:author="A" w:date="2026-10-22T00:00:00Z"><w:r><w:pict><v:shape id="box">'
    "<v:textbox><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></v:textbox>"
    "</v:shape></w:pict></w:r></w:moveTo></mc:Choice><mc:Fallback/></mc:AlternateContent></w:p>"
)

TXBX = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent"


def test_a_text_box_below_a_wildcard_node_is_still_a_wildcard_tree():
    """Why the islands stay (CR-012 section 12.9): below a wildcard node every element is a wildcard.

    An element no class binds (a VML-namespace element the schema does not
    declare), and a branch's ``w:moveFrom`` or ``w:moveTo``, which docx4j
    keeps a DOM element, are such nodes; a text box below either is a wildcard
    ``w:txbxContent``. No corpus document has one.
    """
    for fragment in (BELOW_A_WILDCARD, IN_A_BRANCHS_MOVE):
        assert TXBX in wildcards(wml(fragment))


def test_a_text_box_below_a_wildcard_node_is_read_through_an_island():
    """The text model, the listing and Accept All reach it through a typed island, as before Phase C."""
    import functools

    from docx4j_py import create_package
    from docx4j_py.model import revisions
    from docx4j_py.model.revisions import islands
    from docx4j_py.model.text import extract, plain_text

    made: list = []
    original = islands.Island.__init__

    @functools.wraps(original)
    def counting(self, *args, **kwargs):
        made.append(self)
        original(self, *args, **kwargs)

    pkg = create_package()
    pkg.main_document_part.contents.body.content.insert(0, wml(BELOW_A_WILDCARD))
    islands.Island.__init__ = counting
    try:
        assert plain_text(extract(pkg)) == "before\nboxed\n"
        assert [(change.type, change.text) for change in revisions.tracked_changes(pkg)] == [("Added", "boxed")]
        assert revisions.accept_all(pkg) == 1
        assert revisions.tracked_changes(pkg) == []
    finally:
        islands.Island.__init__ = original
    assert made and all(isinstance(island.typed, CTTxbxContent) for island in made)
