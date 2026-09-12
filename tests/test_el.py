"""``el`` covers every element name the metadata knows, CR-001 section 6.2.

The generator (``codegen/generate_el.py``) derives the (scope, element name) ->
class table by walking the ``XmlContext`` metadata of every generated class.
This test derives it again, independently, and checks that nothing was lost on
the way: every element name the metadata reaches has an ``el`` entry, every
entry builds its class, and every class serialises under the name the entry
claims.

    .venv-fork/bin/python -m pytest tests/test_el.py -q
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "codegen"))

from docx4j_py import wml
from docx4j_py.child import Child
from docx4j_py.el_index import EL_MODULES
from docx4j_py.namespaces import WML_NS
from docx4j_py.runtime import context
from docx4j_py.wml import el

TABLES = ROOT / "codegen" / "el_tables"

# Counts recorded in CR-001 section 14. A change here is a change to the public
# surface of `el` and has to be a deliberate one.
WML_ELEMENT_NAMES = 604
WML_SCOPED_NAMES = 135
WML_KEYWORD_RENAMES = 7
WML_COLLISIONS = 83


@pytest.fixture(scope="module")
def wml_table() -> dict:
    return json.loads((TABLES / "wml.json").read_text(encoding="utf-8"))


# ---------------------------------------------------------------------------
# coverage
# ---------------------------------------------------------------------------


def reachable_element_names(namespace: str) -> set[str]:
    """Every element name of `namespace` the class metadata mentions.

    Derived here the long way round, from the live metadata, so that the test
    does not simply read back what the generator wrote: the element vars of
    every class, the ``choices`` of every compound field, and every class that
    is itself a global element declaration.
    """
    ctx = context()
    prefix = f"{{{namespace}}}"
    names: set[str] = set()
    for cls in list(vars(wml).values()):
        if not isinstance(cls, type) or not issubclass(cls, Child):
            continue
        try:
            meta = ctx.build(cls)
        except Exception:  # noqa: BLE001 - the inkml defect, see CR-001 s.14
            continue
        # A global element declaration is the one thing `Meta` says outright:
        # it carries a namespace as well as a name. `XmlMeta.qname` cannot be
        # used here --- xsdata caches a *type* class's qname against whatever
        # parent namespace it first met it in, so it reads `CT_PPr` before a
        # parse and `{...}CT_PPr` after one (see docx4j_py.traversal).
        own = vars(cls).get("Meta")
        if own is not None and getattr(own, "namespace", None) == namespace:
            local = getattr(own, "name", None)
            if local:
                names.add(f"{prefix}{local}")
        for vars_ in meta.elements.values():
            for var in vars_:
                if var.qname.startswith(prefix):
                    names.add(var.qname)
        for choice in meta.choices:
            names.update(q for q in choice.elements if q.startswith(prefix))
    return names


def test_el_covers_every_reachable_wml_element_name(wml_table):
    """Every element name with a class has an entry; the one without is listed.

    ``w:uniqueTag`` is ``xsd:base64Binary``, so xsdata generated no class for it
    and there is nothing for ``el`` to build: the parent holds the bytes
    directly. It is recorded in the table rather than silently absent.
    """
    reachable = reachable_element_names(WML_NS)
    simple = {f"{{{WML_NS}}}{n}" for n in wml_table["simple_type_elements"]}
    missing = sorted(q for q in reachable if q not in el.QNAME_TO_CLASS)
    assert missing == sorted(simple)
    assert wml_table["simple_type_elements"] == {"uniqueTag": ["bytes"]}
    assert len(el.QNAME_TO_CLASS) >= len(reachable) - len(simple)


def test_wml_element_count_is_the_recorded_one(wml_table):
    assert wml_table["counts"]["element_names"] == WML_ELEMENT_NAMES
    assert len(el.ELEMENTS) == WML_ELEMENT_NAMES
    assert len(el.QNAME_TO_CLASS) == WML_ELEMENT_NAMES


def test_scoped_and_renamed_counts_are_the_recorded_ones(wml_table):
    counts = wml_table["counts"]
    assert counts["scoped_names"] == WML_SCOPED_NAMES
    assert counts["keyword_renames"] == WML_KEYWORD_RENAMES
    assert counts["collisions"] == WML_COLLISIONS
    # every callable in __all__ is a name or a scoped name, plus the 3 tables
    callables = [n for n in el.__all__ if callable(getattr(el, n))]
    assert len(callables) == WML_ELEMENT_NAMES + WML_SCOPED_NAMES


def test_every_generated_namespace_has_an_el_module():
    from importlib import import_module

    assert len(EL_MODULES) >= 50
    for uri, module in EL_MODULES.items():
        mod = import_module(module)
        assert mod.__NAMESPACE__ == uri
        assert set(mod.QNAME_TO_CLASS) == {f"{{{uri}}}{n}" for n in mod.ELEMENTS}


# ---------------------------------------------------------------------------
# what the functions build
# ---------------------------------------------------------------------------


def test_every_entry_builds_its_class_and_is_a_child():
    for name, clazz in el.ELEMENTS.items():
        built = getattr(el, el_name(name))()
        assert isinstance(built, clazz), name
        assert isinstance(built, Child), name


def el_name(local: str) -> str:
    """The function name of an element, mirroring the generator's rule."""
    for entry, value in el.ELEMENTS.items():
        if entry == local:
            break
    import builtins
    import keyword

    reserved = set(keyword.kwlist) | set(keyword.softkwlist) | set(dir(builtins))
    return f"{local}_" if local in reserved else local


def test_every_entry_serialises_under_its_qualified_name():
    """The name a class writes is the name the table filed it under.

    The interesting half is the 72 intermediate choice classes (CR-001 section
    13.3), which have no element name of their own: ``el.t()`` is an ``RT``, and
    an ``RT`` only becomes ``w:t`` because the table says so. ``to_xml`` uses
    the same table, so this pins the pair.
    """
    from docx4j_py.fragments import to_xml
    from docx4j_py.traversal import element_name

    checked = 0
    for local, clazz in el.ELEMENTS.items():
        if el.CLASS_TO_QNAME.get(clazz) is None:
            continue  # the class serves several element names; to_xml needs name=
        obj = getattr(el, el_name(local))()
        assert element_name(obj) == f"{{{WML_NS}}}{local}", local
        assert to_xml(obj).startswith(f"<w:{local} "), local
        checked += 1
    assert checked > 300


@pytest.mark.parametrize(
    ("name", "local"),
    [
        ("p", "p"),
        ("r", "r"),
        ("t", "t"),
        ("tbl", "tbl"),
        ("tc", "tc"),
        ("pPr", "pPr"),
        ("rPr", "rPr"),
        ("document", "document"),
        ("styles", "styles"),
        ("numbering", "numbering"),
        ("settings", "settings"),
        ("fonts", "fonts"),
        ("hdr", "hdr"),
        ("ftr", "ftr"),
        ("comments", "comments"),
        ("footnotes", "footnotes"),
        ("endnotes", "endnotes"),
    ],
)
def test_the_names_everyone_uses(name, local):
    from docx4j_py.fragments import to_xml

    obj = getattr(el, name)()
    assert to_xml(obj).startswith(f"<w:{local} ")


def test_part_roots_map_to_the_root_classes():
    assert el.document().__class__ is wml.Document
    assert el.styles().__class__ is wml.Styles
    assert el.numbering().__class__ is wml.Numbering
    assert el.fonts().__class__ is wml.Fonts
    assert el.hdr().__class__ is wml.Hdr
    assert el.ftr().__class__ is wml.Ftr
    # `Settings` is the class of the global element `w:settings`; `CTSettings`
    # is the type it extends, and the element class wins (CR-001 section 13.3)
    assert el.settings().__class__ is wml.Settings
    assert isinstance(el.settings(), wml.CTSettings)
    assert el.comments().__class__ is wml.Comments


def test_element_specific_subclasses_win_their_names():
    """Decided question 2: keep xsdata's element-specific subclasses."""
    assert el.ins().__class__ is wml.RunIns
    assert el.del_().__class__ is wml.RunDel
    assert el.t().__class__ is wml.RT
    assert el.delText().__class__ is wml.DelText
    assert isinstance(el.t(), wml.Text)
    # `find(root, CTTrackChange)` must reach both, which decided question 2
    # promised and `tests/test_traversal.py` checks on a real document
    assert isinstance(el.ins(), wml.CTTrackChange)
    assert isinstance(el.del_(), wml.CTTrackChange)
    assert isinstance(el.ins(), wml.CTMarkup)


# ---------------------------------------------------------------------------
# the naming rules
# ---------------------------------------------------------------------------


def test_python_keywords_and_builtins_take_a_trailing_underscore(wml_table):
    assert wml_table["keyword_renames"] == {
        "del": "del_",
        "dir": "dir_",
        "format": "format_",
        "id": "id_",
        "next": "next_",
        "object": "object_",
        "type": "type_",
    }
    assert not hasattr(el, "del")
    assert callable(el.del_)
    # `name` is neither a keyword nor a builtin, so it keeps its own name
    assert callable(el.name)


def test_scope_qualified_names_for_the_documented_collisions():
    assert el.sdt().__class__ is wml.SdtBlock
    assert el.sdt_run().__class__ is wml.SdtRun
    assert el.sdt_row().__class__ is wml.CTSdtRow
    assert el.sdt_cell().__class__ is wml.CTSdtCell

    assert el.customXml().__class__ is wml.CTCustomXmlBlock
    assert el.custom_xml_run().__class__ is wml.CTCustomXmlRun
    assert el.custom_xml_row().__class__ is wml.CTCustomXmlRow
    assert el.custom_xml_cell().__class__ is wml.CTCustomXmlCell

    assert el.t_math().__class__ is wml.T
    assert el.ins_math().__class__ is wml.Ins1
    assert el.ins_ctrl_pr().__class__ is wml.Ins2
    assert el.del_math().__class__ is wml.Del1
    assert el.del_track_change().__class__ is wml.CTTrackChange

    assert el.hyperlink().__class__ is wml.PHyperlink
    assert el.simple_field_hyperlink().__class__ is wml.CtSimpleFieldHyperlink


def test_every_collision_is_in_the_committed_table(wml_table):
    """Nothing resolves silently: a name with more than one class is recorded."""
    for local, entry in wml_table["collisions"].items():
        assert getattr(el, el_name(local)).__doc__.startswith(
            f"``{{{WML_NS}}}{local}``"
        )
        for scoped in entry["scoped"]:
            assert callable(getattr(el, scoped["name"])), scoped
            assert scoped["scopes"], scoped
        assert entry["default"]["chosen_by"] in {
            "override",
            "global element declaration",
            "most reachable scope",
        }


def test_the_overrides_a_human_set_are_still_applied(wml_table):
    overrides = wml_table["overrides"]
    assert overrides["defaults"] == {
        "hyperlink": "PHyperlink",
        "name": "StyleName",
        "spacing": "CtPprBaseSpacing",
        "customXml": "CTCustomXmlBlock",
    }
    for local, want in overrides["defaults"].items():
        assert el.ELEMENTS[local].__name__ == want
    for key, name in overrides["aliases"].items():
        local, _, cls_name = key.partition(":")
        assert getattr(el, name)().__class__.__name__ == cls_name, (local, name)


# ---------------------------------------------------------------------------
# positional text, decided question 6
# ---------------------------------------------------------------------------


def test_text_carrying_classes_take_the_text_positionally():
    assert el.t("Hello").value == "Hello"
    assert el.delText("gone").value == "gone"
    assert el.instrText(" PAGE ").value == " PAGE "


@pytest.mark.parametrize(
    ("text", "preserve"),
    [
        ("Hello", False),
        (" leading", True),
        ("trailing ", True),
        ("two  spaces", True),
        ("a\tb", False),  # one whitespace character, neither leading nor trailing
        ("a \t b", True),
        ("", False),
        ("one space here", False),
    ],
)
def test_xml_space_is_set_when_the_whitespace_needs_it(text, preserve):
    assert (el.t(text).space == "preserve") is preserve


def test_only_text_carrying_classes_take_a_positional_argument():
    with pytest.raises(TypeError):
        el.p("Hello")  # a paragraph is keywords only
    with pytest.raises(TypeError):
        el.t("x", value="y")


def test_keyword_arguments_go_straight_through():
    run = el.r(content=[el.t("x")])
    assert run.content[0].value == "x"
    assert el.br(type_value="page").type_value == "page"
    with pytest.raises(TypeError):
        el.p(no_such_field=1)


def test_children_are_linked_when_placed():
    """Everything `el` builds is a Child; ChildList links it when placed."""
    text = el.t("x")
    run = el.r()
    run.content.append(text)
    assert text.parent is run
    paragraph = el.p(content=[run])
    assert run.parent is paragraph
