"""What a save writes for an integer attribute: CR-021 sections 4.3 to 4.6 (D3, D5, D6), row by row.

``tests/content/test_integer_values.py`` holds the rows over docx4j CR-065's probe
and Word's saves; this file holds each rule on its own: the typed rows in
``w``, ``w14`` and ``w15`` (Word's reading), the rows elsewhere (today's output:
Python's ``int()`` where it takes the value, the value as read where it does
not), the union fields' rows in the render, the empty rsid, the canon's lxml
mirror of all of them (Q5), and section 2.2's forms, which the fork's stage 14
keeps as their text once the engine reads integers exactly (phase 0).
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

from docx4j_py.model.values import integer
from docx4j_py.openpackaging import normalise
from docx4j_py.openpackaging.normalise import (
    INTEGER_NAMESPACES,
    UNION_INTEGER_ATTRIBUTES,
    integer_attributes,
    normalised_tree,
    rendered_rows,
    rsid_attributes,
)

ROOT = Path(__file__).resolve().parents[2]
INTEGERS = ROOT / "tests" / "fixtures" / "integers"
SAMPLES = ROOT / "samples"

pytestmark = pytest.mark.filterwarnings("ignore::docx4j_xsdata.exceptions.ConverterWarning")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _w(local: str) -> str:
    return f"{{{W}}}{local}"


def _exact_integers() -> bool:
    """Whether the engine reads integers exactly: the fork's stage 14 in ``runtime.EXACT_TYPES`` (phase 0)."""
    from docx4j_py.runtime import EXACT_TYPES

    return int in EXACT_TYPES


# ---------------------------------------------------------------------------
# which attributes
# ---------------------------------------------------------------------------


def test_integer_attributes() -> None:
    from docx4j_py.dml.diagram import CTForEach
    from docx4j_py.dml.main import CTPositiveSize2D
    from docx4j_py.wml import CtPprBaseSpacing, CtSectPrPgMar

    spacing = {var.name: word for var, word in integer_attributes(CtPprBaseSpacing)}
    assert spacing == {
        "before": True, "before_lines": True, "after": True, "after_lines": True, "line": True,
    }
    assert {var.name for var, _word in integer_attributes(CtSectPrPgMar)} == {
        "top", "right", "bottom", "left", "header", "footer", "gutter",
    }
    assert {var.name: word for var, word in integer_attributes(CTPositiveSize2D)} == {"cx": False, "cy": False}
    tokens = {var.name: (word, bool(var.tokens)) for var, word in integer_attributes(CTForEach)}
    assert tokens == {"st": (False, True), "cnt": (False, True), "step": (False, True)}
    assert integer_attributes(CtPprBaseSpacing) is integer_attributes(CtPprBaseSpacing)  # cached


def test_rsid_attributes() -> None:
    from docx4j_py.wml import P, R, SectPr, Tr

    names = {cls.__name__: sorted(var.qname.split("}")[1] for var in rsid_attributes(cls)) for cls in (P, R, SectPr, Tr)}
    assert names == {
        "P": ["rsidDel", "rsidP", "rsidR", "rsidRDefault", "rsidRPr"],
        "R": ["rsidDel", "rsidR", "rsidRPr"],
        "SectPr": ["rsidDel", "rsidR", "rsidRPr", "rsidSect"],
        "Tr": ["rsidDel", "rsidR", "rsidRPr", "rsidTr"],
    }


# ---------------------------------------------------------------------------
# the typed rows
# ---------------------------------------------------------------------------


def _paragraph(**spacing: Any) -> Any:
    from docx4j_py.child import link_parents
    from docx4j_py.wml import el

    p = el.p(p_pr=el.pPr(spacing=el.spacing(**spacing)))
    link_parents(p)
    return p


def test_the_word_rows() -> None:
    """A kept decimal, ``""`` and a refused form, written as Word reads them; the tree restored after."""
    p = _paragraph(before="120.5", after="", line="abc", before_lines=240)
    with normalised_tree(p, part="/word/document.xml") as records:
        assert (p.p_pr.spacing.before, p.p_pr.spacing.after, p.p_pr.spacing.line) == (120, 0, None)
    assert [(r.path, r.before, r.after) for r in records] == [
        ("w:p/w:pPr/w:spacing/@w:before", "120.5", "120"),
        ("w:p/w:pPr/w:spacing/@w:after", "", "0"),
        ("w:p/w:pPr/w:spacing/@w:line", "abc", None),
    ]
    assert (p.p_pr.spacing.before, p.p_pr.spacing.after, p.p_pr.spacing.line) == ("120.5", "", "abc")
    assert p.p_pr.spacing.before_lines == 240


def test_a_value_word_reads_as_written_is_not_a_row() -> None:
    """``"240"`` held as a ``str`` (code's) is Word's form already; an ``int`` is the serialiser's."""
    p = _paragraph(before="240", after=120)
    with normalised_tree(p) as records:
        assert records == []


def test_the_rows_outside_word_s_namespaces_keep_today_s_output() -> None:
    """D5: a kept value is written as Python's ``int()`` binds it, where it can; as read where it cannot."""
    from docx4j_py.child import link_parents
    from docx4j_py.dml.diagram import CTForEach
    from docx4j_py.dml.main import CTPositiveSize2D
    from docx4j_py.dml.wordprocessing_drawing import Inline

    inline = Inline(extent=CTPositiveSize2D(cx=" 5", cy="1.5"))
    link_parents(inline)
    with normalised_tree(inline) as records:
        assert (inline.extent.cx, inline.extent.cy) == (5, "1.5")
    assert [(r.before, r.after) for r in records] == [(" 5", "5")]
    assert (inline.extent.cx, inline.extent.cy) == (" 5", "1.5")

    each = CTForEach(st="1 2_0", cnt="2 x")
    with normalised_tree(each) as records:
        assert (each.st, each.cnt) == ("1 20", "2 x")
    assert [(r.before, r.after) for r in records] == [("1 2_0", "1 20")]


def test_the_union_rows_in_the_render() -> None:
    """D5: in a union field holding an integer, a value ``int()`` takes is written as it, by name in the render."""
    root = etree.fromstring(
        f'<a:root xmlns:a="{A}"><a:pt x=" 240" y="guide"/><a:pt x="240" y="1_000"/><a:rect l="1.5" t="+7"/></a:root>'
    )
    records = rendered_rows(root, "/word/document.xml")
    assert [(r.path, r.before, r.after) for r in records] == [
        ("a:root/a:pt[1]/@x", " 240", "240"),
        ("a:root/a:pt[2]/@y", "1_000", "1000"),
        ("a:root/a:rect/@t", "+7", "7"),
    ]
    assert [p.get("x") for p in root] == ["240", "240", None]


def test_the_union_table_is_the_model_s() -> None:
    """:data:`UNION_INTEGER_ATTRIBUTES` is every (element, attribute) whose field is a union holding ``int`` (35 fields)."""
    import dataclasses
    import inspect
    from importlib import import_module

    from docx4j_py.el_index import EL_MODULES
    from docx4j_py.runtime import context

    def unions(cls: type) -> list[Any]:
        return [var for var in context().build(cls).get_attribute_vars() if int in var.types and tuple(var.types) != (int,)]

    bound, _holders, broken = normalise._bound()
    derived: dict[str, set[str]] = {}
    for element, classes in bound.items():
        for cls in classes - broken:
            for var in unions(cls):
                derived.setdefault(element, set()).add(var.qname)
    assert {element: set(names) for element, names in UNION_INTEGER_ATTRIBUTES.items()} == derived
    fields = set()
    for module_name in sorted(set(EL_MODULES.values())):
        for cls in vars(import_module(module_name.removesuffix(".el"))).values():
            if inspect.isclass(cls) and dataclasses.is_dataclass(cls) and cls not in broken:
                fields.update((cls, var.name) for var in unions(cls))
    assert len(fields) == 35  # CR-021 section 4.5; one, msink's CT_CtxNode/@semanticType, no element binds
    assert not any(normalise._namespace(var) in INTEGER_NAMESPACES for names in derived.values() for var in names)


# ---------------------------------------------------------------------------
# the empty rsid (D6)
# ---------------------------------------------------------------------------


def _docx_with(document: bytes, source: Path = SAMPLES / "2010-sample1.docx") -> bytes:
    """`source` with its main part replaced."""
    out = io.BytesIO()
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, document if item.filename == "word/document.xml" else zin.read(item.filename))
    return out.getvalue()


def test_an_empty_rsid_is_dropped_where_nothing_was_kept() -> None:
    """SSRS's ``w:rsidR=""`` binds without a record; the source's empty value sends the part through the rows."""
    from docx4j_py.openpackaging.api import load

    source = zipfile.ZipFile(SAMPLES / "2010-sample1.docx").read("word/document.xml")
    assert b'w:rsidR=""' not in source
    edited = source.replace(b"<w:p ", b'<w:p w:rsidR="" ', 1)
    package = load(io.BytesIO(_docx_with(edited)))
    part = package.main_document_part
    _ = part.contents
    assert part.kept == [] and part._empty_value is True
    first = part.contents.body.content[0]
    assert first.rsid_r == ""
    xml = zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml")
    assert b'w:rsidR=""' not in xml
    assert [(r.path, r.before, r.after) for r in package.last_save.normalised] == [
        ("w:document/w:body/w:p[1]/@w:rsidR", "", None)
    ]
    assert first.rsid_r == ""  # the tree keeps what was read
    assert b'w:rsidR=""' in zipfile.ZipFile(io.BytesIO(package.save(normalise=False))).read("word/document.xml")


def test_a_part_with_nothing_odd_is_not_walked() -> None:
    from docx4j_py.openpackaging.api import load

    package = load(SAMPLES / "Symbols.docx")
    part = package.main_document_part
    _ = part.contents
    assert part.kept == [] and part._empty_value is False
    package.save()
    assert package.last_save.normalised == []


def test_a_value_set_in_code_keeps_today_s_form() -> None:
    """As for the on/off rows: a ``str`` set in code, in a part whose parse kept nothing, is written as set."""
    from docx4j_py.openpackaging.api import create_package
    from docx4j_py.wml import el

    package = create_package()
    body = package.main_document_part.contents.body
    body.content.append(el.p(p_pr=el.pPr(spacing=el.spacing(before="120.5"))))
    xml = zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml")
    assert b'w:before="120.5"' in xml and package.last_save.normalised == []


# ---------------------------------------------------------------------------
# the canon's mirror (Q5)
# ---------------------------------------------------------------------------


def test_the_canon_s_integer_tables() -> None:
    """Every class that declares the attribute types it ``int``, alike; the three pairs that disagree, by holder."""
    table, by_holder = normalise._integer_tables()
    in_w = [(e, a) for (e, a), kind in table.items() if e.startswith(f"{{{W}}}") and a.startswith(f"{{{W}}}")]
    assert len(in_w) == 217 and all(table[pair] == "word" for pair in in_w)  # 192 declared by every binding, 25 not
    assert by_holder == {
        (_w("rPr"), _w("sz"), _w("val")): "word",
        (_w("ddList"), _w("default"), _w("val")): "word",
        (_w("lvl"), _w("start"), _w("val")): "word",
    }
    for ambiguous in ("sz", "default", "start"):
        assert (_w(ambiguous), _w("val")) not in table
    assert table[(_w("bottom"), _w("sz"))] == "word"  # a border's; w:tblCellMar's w:bottom has no w:sz
    assert table[(f"{{{A}}}off", "x")] == "python"
    assert table[("{http://schemas.openxmlformats.org/drawingml/2006/diagram}forEach", "st")] == "tokens"


@pytest.mark.parametrize(
    ("value", "kind", "exact", "written"),
    [
        ("720.5", "word", False, "720"), ("", "word", False, "0"), ("abc", "word", False, None),
        ("+240", "word", False, "240"), (" 240 ", "word", False, "240"), ("1_000", "word", False, "1000"),
        (" 240 ", "word", True, None), ("1_000", "word", True, None), ("+240", "word", True, "240"),
        ("1.5", "python", False, "1.5"), (" 5", "python", True, "5"), ("1 2_0", "tokens", True, "1 20"),
        ("1 x", "tokens", False, "1 x"),
    ],
)
def test_the_canon_s_normal_form(value: str, kind: str, exact: bool, written: str | None) -> None:
    """Before stage 14 a value ``int()`` takes binds, and the serialiser writes it; after, Word's rule decides."""
    assert normalise._integer_normal_form(value, kind, exact) == written


def test_normalise_tree_finds_by_name_and_by_holder() -> None:
    root = etree.fromstring(
        f'<w:document xmlns:w="{W}"><w:body><w:p w:rsidR="" w:rsidRDefault="00AB12CD"><w:pPr>'
        '<w:spacing w:before="120.5" w:after="+240"/><w:framePr w:w="10.5"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="22.5"/></w:rPr></w:r></w:p>'
        '<w:p><w:pPr><w:framePr/></w:pPr></w:p></w:body></w:document>'
    )
    rows = normalise.normalise_tree(root)
    assert [(path.rsplit("/", 1)[-1], attribute.split("}")[1], before, after) for path, attribute, before, after in rows] == [
        ("w:p[1]", "rsidR", "", None),
        ("w:spacing", "before", "120.5", "120"),
        ("w:spacing", "after", "+240", "240"),
        ("w:framePr", "w", "10.5", "10"),
        ("w:sz", "val", "22.5", "22"),
    ]


# ---------------------------------------------------------------------------
# section 2.2: the forms the fork's stage 14 keeps (phase 0)
# ---------------------------------------------------------------------------

#: Forms Python's ``int()`` binds and Word refuses, and the reason the fork's stage 14 keeps each: XSD's
#: whitespace alone is kept on purpose (``"exact"``); underscores, other scripts' digits and Unicode
#: whitespace no XSD rule converts (``"invalid"``, with the warning). Each reads None and is dropped at save.
#: Until phase 0 turns ``int`` on in ``runtime.EXACT_TYPES`` they bind as 240 (section 2.2).
STAGE_14_FORMS = [
    (" 240 ", "exact"),
    ("\t240\n", "exact"),
    ("1_000", "invalid"),
    ("\u0662\u0664\u0660", "invalid"),  # Arabic-Indic digits
    ("\uff12\uff14\uff10", "invalid"),  # fullwidth digits
    ("\u00a0240", "invalid"),  # a no-break space
]


def _a1_with(form: str) -> bytes:
    path = INTEGERS / "cr065-A1-decimals.docx"
    document = zipfile.ZipFile(path).read("word/document.xml")
    # character references, so that XML's attribute normalisation leaves the tab, newline and no-break space alone
    xml_form = "".join(f"&#{ord(c)};" if c in "\t\n\u00a0" else c for c in form).encode("utf-8")
    assert document.count(b'w:before="120.5"') == 1
    return _docx_with(document.replace(b'w:before="120.5"', b'w:before="' + xml_form + b'"'), path)


@pytest.mark.skipif(not _exact_integers(), reason="phase 0: int is not in runtime.EXACT_TYPES until docx4j-xsdata 26.2.4")
@pytest.mark.parametrize(("form", "reason"), STAGE_14_FORMS)
def test_a_form_word_refuses_is_kept_read_none_and_dropped(form: str, reason: str) -> None:
    import warnings

    from docx4j_py.openpackaging.api import load

    package = load(io.BytesIO(_a1_with(form)))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        body = package.main_document_part.contents.body
    # (A1 holds the decimals too, each kept and warned about: only the form's own count here)
    warned = [w for w in caught if f"`{form}`" in str(w.message)]
    assert len(warned) == (1 if reason == "invalid" else 0), [str(w.message) for w in caught]
    spacing = body.content[1].p_pr.spacing
    assert spacing.before == form
    kept = [k for _part, k in package.kept if k.text == form]
    assert [(k.field_name, k.reason) for k in kept] == [("before", reason)]
    assert integer(spacing.before) is None
    assert package.body.paragraphs[1].space_before == 0.0  # unset
    xml = zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml")
    first = etree.fromstring(xml).find(f"{_w('body')}/{_w('p')}[2]/{_w('pPr')}/{_w('spacing')}")
    assert first.get(_w("before")) is None
    assert [(r.before, r.after) for r in package.last_save.normalised if r.before == form] == [(form, None)]
    assert spacing.before == form  # the tree keeps what was read


@pytest.mark.skipif(not _exact_integers(), reason="phase 0: int is not in runtime.EXACT_TYPES until docx4j-xsdata 26.2.4")
def test_outside_word_s_namespaces_the_output_is_today_s() -> None:
    """D5 with the exact parse: a padded DrawingML integer is kept, and written as ``int()`` bound it before."""
    from docx4j_py.dml.main import CTPositiveSize2D
    from docx4j_py.fragments import wml  # noqa: F401 - the model loaded
    from docx4j_py.runtime import context, parser

    xml = f'<a:ext xmlns:a="{A}" cx=" 5" cy="7"/>'
    parsed = parser(lenient=True, skipped_report=True).from_string(xml, CTPositiveSize2D)
    assert parsed.cx == " 5" and parsed.cy == 7
    with normalised_tree(parsed) as records:
        pass
    assert [(r.before, r.after) for r in records] == [(" 5", "5")]
    assert context() is not None


@pytest.mark.skipif(not _exact_integers(), reason="phase 0: int is not in runtime.EXACT_TYPES until docx4j-xsdata 26.2.4")
@pytest.mark.parametrize("value", [" 240", "240 ", "1_000", "\u0662\u0664\u0660", "\u00a0240", "+7", "1.5", "guide", ""])
def test_the_union_fields_write_what_26_2_3_wrote(value: str) -> None:
    """D5 over every union field an element binds: stage 14's parse and the render's rows give 26.2.3's output.

    26.2.3's parse is ``exact_types={bool, Enum}`` (Python's ``int()``); the
    element is rendered under its own name, as under its parent.
    """
    from enum import Enum

    from docx4j_xsdata.formats.dataclass.parsers import XmlParser

    from docx4j_py.runtime import context, parser_config, tree_serializer

    bound, _holders, broken = normalise._bound()
    checked = 0
    text = "".join(f"&#{ord(c)};" if c == "\u00a0" else c for c in value)
    for element, names in UNION_INTEGER_ATTRIBUTES.items():
        namespace, local = element[1:].split("}")
        for cls in bound[element] - broken:
            for attribute in names:
                written = attribute if not attribute.startswith("{") else "x:" + attribute.split("}")[1]
                xml = f'<x:{local} xmlns:x="{namespace}" {written}="{text}"/>'
                outputs = []
                for exact in (frozenset({bool, Enum}), frozenset({bool, Enum, int})):
                    parsed = XmlParser(config=parser_config(exact_types=exact), context=context()).from_string(xml, cls)
                    root = tree_serializer().render(parsed, {"x": namespace}).getroot()
                    root.tag = element
                    if int in exact:
                        rendered_rows(root)
                    outputs.append(root.get(attribute))
                assert outputs[0] == outputs[1], (element, attribute, value, outputs)
                checked += 1
    assert checked == 42


@pytest.mark.parametrize("form", ["+240", "0240"])
def test_a_sign_or_a_leading_zero_reads_240_either_way(form: str) -> None:
    """``"+240"`` and ``"0240"`` are xs:integer forms: they bind as 240 before stage 14 and after (section 2.2's last row)."""
    from docx4j_py.openpackaging.api import load

    package = load(io.BytesIO(_a1_with(form)))
    assert package.main_document_part.contents.body.content[1].p_pr.spacing.before == 240
