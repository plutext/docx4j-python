"""Tests for docx4j_py.child: the parent slot, ChildList, link_parents, deep_copy.

    .venv-fork/bin/python -m pytest tests/ -q

Two halves. The first works on small hand made dataclasses shaped exactly like
the generated ones, so a failure points at ``child.py`` and nothing else. The
second works on the real generated WordprocessingML package and on parsed
sample parts, which is where the interaction with xsdata's metadata,
``@dataclass(slots=True, kw_only=True)`` and the serialiser actually shows up.
"""

from __future__ import annotations

import copy
import dataclasses
import pickle
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py.child import (
    Child,
    ChildList,
    deep_copy,
    iter_tree,
    link_parents,
)

# ---------------------------------------------------------------------------
# hand made models, shaped like the generated ones
# ---------------------------------------------------------------------------


@dataclass(slots=True, kw_only=True)
class Text(Child):
    value: None | str = None


@dataclass(slots=True, kw_only=True)
class RPr(Child):
    bold: None | bool = None


@dataclass(slots=True, kw_only=True)
class R(Child):
    r_pr: None | RPr = None
    content: list = field(default_factory=ChildList)


@dataclass(slots=True, kw_only=True)
class PPr(Child):
    style: None | str = None


@dataclass(slots=True, kw_only=True)
class P(Child):
    p_pr: None | PPr = None
    content: list = field(default_factory=ChildList)


def sample_paragraph() -> P:
    """A paragraph with a run with text, built through the constructors."""
    return P(p_pr=PPr(style="Heading1"), content=[R(content=[Text(value="Hello")])])


# ---------------------------------------------------------------------------
# the base class mechanics
# ---------------------------------------------------------------------------


class TestChildBase:
    def test_parent_is_a_slot_not_a_field(self):
        assert Child.__slots__ == ("parent",)
        assert [f.name for f in dataclasses.fields(P)] == ["p_pr", "content"]

    def test_slots_subclass_of_a_slotted_base(self):
        """dataclass(slots=True) does not redeclare the inherited slot."""
        assert P.__slots__ == ("p_pr", "content")
        assert not hasattr(P(), "__dict__")
        with pytest.raises(AttributeError):
            P().nope = 1

    def test_post_init_on_the_base_is_invoked(self):
        """dataclasses calls an inherited __post_init__; the slot starts None."""
        assert P().parent is None
        assert P().get_parent() is None

    def test_set_and_get_parent(self):
        p, r = P(), R()
        r.set_parent(p)
        assert r.get_parent() is p
        r.set_parent(None)
        assert r.get_parent() is None

    def test_equality_and_repr_ignore_the_parent(self):
        a, b = Text(value="x"), Text(value="x")
        b.set_parent(P())
        assert a == b
        assert repr(a) == repr(b) == "Text(value='x')"

    def test_a_plain_child_subclass_without_dataclass(self):
        class Bare(Child):
            __slots__ = ()

        assert Bare().get_parent() is None


# ---------------------------------------------------------------------------
# ChildList
# ---------------------------------------------------------------------------


class TestChildList:
    def test_default_factory_is_bound_by_post_init(self):
        p = P()
        assert isinstance(p.content, ChildList)
        assert p.content.owner is p

    def test_append_insert_extend_adopt(self):
        p = P()
        a, b, c = R(), R(), R()
        p.content.append(a)
        p.content.insert(0, b)
        p.content.extend([c])
        assert [x.get_parent() for x in (a, b, c)] == [p, p, p]

    def test_iadd_adopts(self):
        p, r = P(), R()
        p.content += [r]
        assert r.get_parent() is p
        assert isinstance(p.content, ChildList)

    def test_setitem_by_index_and_slice(self):
        p = P()
        a, b, c = R(), R(), R()
        p.content.append(a)
        p.content[0] = b
        assert b.get_parent() is p
        assert a.get_parent() is None
        p.content[0:1] = [c]
        assert c.get_parent() is p
        assert b.get_parent() is None

    def test_removal_orphans(self):
        p = P()
        a, b, c, d = R(), R(), R(), R()
        p.content.extend([a, b, c, d])
        p.content.remove(a)
        assert p.content.pop() is d
        del p.content[0]
        p.content.clear()
        assert [x.get_parent() for x in (a, b, c, d)] == [None] * 4

    def test_removal_can_keep_the_parent(self, monkeypatch):
        monkeypatch.setattr(ChildList, "orphan_on_remove", False)
        p, r = P(), R()
        p.content.append(r)
        p.content.remove(r)
        assert r.get_parent() is p

    def test_a_plain_list_in_the_constructor_is_promoted(self):
        r = R()
        p = P(content=[r])
        assert isinstance(p.content, ChildList)
        assert p.content.owner is p
        assert r.get_parent() is p

    def test_non_child_items_are_left_alone(self):
        p = P()
        p.content.extend(["text", 3, None])
        assert list(p.content) == ["text", 3, None]

    def test_pickle_drops_the_owner_and_link_parents_restores_it(self):
        p = sample_paragraph()
        back = pickle.loads(pickle.dumps(p))
        assert back == p
        assert back.content.owner is None
        link_parents(back)
        assert back.content.owner is back
        assert back.content[0].get_parent() is back


# ---------------------------------------------------------------------------
# link_parents
# ---------------------------------------------------------------------------


def assert_parents_consistent(root) -> int:
    """Every node's parent is the object that holds it; the root's is untouched."""
    count = 0
    stack = [(root, root.get_parent())]
    seen = {id(root)}
    while stack:
        obj, want = stack.pop()
        count += 1
        assert obj.get_parent() is want, f"{type(obj).__name__} has the wrong parent"
        for f in dataclasses.fields(obj):
            value = getattr(obj, f.name, None)
            if isinstance(value, ChildList):
                assert value.owner is obj
            items = value if isinstance(value, (list, tuple)) else (value,)
            for item in items:
                if isinstance(item, Child) and id(item) not in seen:
                    seen.add(id(item))
                    stack.append((item, obj))
    return count


class TestLinkParents:
    def test_wires_single_valued_fields(self):
        p = P()
        p.p_pr = PPr(style="X")
        assert p.p_pr.get_parent() is None
        link_parents(p)
        assert p.p_pr.get_parent() is p

    def test_wires_a_whole_tree(self):
        p = sample_paragraph()
        assert link_parents(p) == 4
        assert_parents_consistent(p)

    def test_leaves_the_root_parent_alone(self):
        outer, p = P(), sample_paragraph()
        p.set_parent(outer)
        link_parents(p)
        assert p.get_parent() is outer

    def test_rebinds_a_replaced_list(self):
        p = sample_paragraph()
        r = R()
        p.content = [r]  # a plain list assigned over the ChildList
        link_parents(p)
        assert r.get_parent() is p

    def test_is_idempotent(self):
        p = sample_paragraph()
        assert link_parents(p) == link_parents(p) == 4

    def test_iter_tree_visits_everything_once(self):
        p = sample_paragraph()
        assert len(list(iter_tree(p))) == 4


# ---------------------------------------------------------------------------
# deep_copy
# ---------------------------------------------------------------------------


class TestDeepCopy:
    def test_copy_equals_the_original_and_shares_nothing(self):
        p = sample_paragraph()
        link_parents(p)
        c = deep_copy(p)
        assert c == p
        assert c is not p
        assert c.content[0] is not p.content[0]

    def test_the_copy_is_detached_and_internally_linked(self):
        p = sample_paragraph()
        link_parents(p)
        c = deep_copy(p)
        assert c.get_parent() is None
        assert_parents_consistent(c)

    def test_copying_a_subtree_does_not_copy_the_document(self):
        p = sample_paragraph()
        link_parents(p)
        run = p.content[0]
        assert run.get_parent() is p
        c = deep_copy(run)
        assert c.get_parent() is None
        assert c == run

    def test_parent_argument_attaches_the_copy(self):
        p = sample_paragraph()
        link_parents(p)
        other = P()
        c = deep_copy(p.content[0], parent=other)
        assert c.get_parent() is other

    def test_copy_deepcopy_behaves_the_same(self):
        p = sample_paragraph()
        link_parents(p)
        c = copy.deepcopy(p)
        assert c == p
        assert c.get_parent() is None
        assert_parents_consistent(c)


# ---------------------------------------------------------------------------
# the generated package
# ---------------------------------------------------------------------------

wml = pytest.importorskip("docx4j_py.wml")


@pytest.fixture(scope="module")
def context():
    from docx4j_xsdata.formats.dataclass.context import XmlContext

    return XmlContext()


@pytest.fixture(scope="module")
def parser(context):
    from docx4j_xsdata.formats.dataclass.parsers import XmlParser
    from docx4j_xsdata.formats.dataclass.parsers.config import ParserConfig
    from docx4j_xsdata.formats.dataclass.parsers.handlers import LxmlEventHandler

    return XmlParser(
        config=ParserConfig(
            fail_on_unknown_properties=False,
            fail_on_unknown_attributes=False,
            fail_on_converter_warnings=False,
        ),
        context=context,
        handler=LxmlEventHandler,
    )


@pytest.fixture(scope="module")
def serializer(context):
    from docx4j_xsdata.formats.dataclass.serializers import XmlSerializer
    from docx4j_xsdata.formats.dataclass.serializers.config import SerializerConfig
    from docx4j_xsdata.formats.dataclass.serializers.writers import LxmlEventWriter

    return XmlSerializer(
        config=SerializerConfig(xml_declaration=True, indent=None),
        context=context,
        writer=LxmlEventWriter,
    )


@pytest.fixture(scope="module")
def document_bytes() -> bytes:
    with zipfile.ZipFile(ROOT / "samples" / "tables.docx") as z:
        return z.read("word/document.xml")


@pytest.fixture(scope="module")
def document(parser, document_bytes):
    return parser.from_bytes(document_bytes, wml.Document)


class TestGeneratedPackage:
    def test_every_generated_class_is_a_child(self):
        klasses = [
            obj
            for name, obj in vars(wml).items()
            if isinstance(obj, type) and dataclasses.is_dataclass(obj)
        ]
        assert len(klasses) > 500
        assert all(issubclass(k, Child) for k in klasses)

    def test_the_generated_shape_is_the_one_the_cr_asked_for(self):
        params = wml.P.__dataclass_params__
        assert params.slots and params.kw_only
        assert wml.P.__mro__[1] is Child
        assert "parent" not in {f.name for f in dataclasses.fields(wml.P)}

    def test_list_fields_use_childlist(self):
        content = next(f for f in dataclasses.fields(wml.P) if f.name == "content")
        assert content.default_factory is ChildList
        assert wml.P().content.owner is not None

    def test_construction_wires_list_children(self):
        r = wml.R()
        p = wml.P(content=[r])
        assert r.get_parent() is p

    def test_parse_then_link(self, document, context):
        n = link_parents(document, context=context)
        assert n > 1000
        assert document.get_parent() is None
        body = document.body
        assert body.get_parent() is document

    def test_every_parsed_node_has_the_right_parent(self, document, context):
        link_parents(document, context=context)
        checked = 0
        for node in iter_tree(document, context=context):
            for name in (f.name for f in dataclasses.fields(node)):
                value = getattr(node, name, None)
                if isinstance(value, ChildList):
                    assert value.owner is node
                items = value if isinstance(value, (list, tuple)) else (value,)
                for item in items:
                    if isinstance(item, Child):
                        assert item.get_parent() is node
                        checked += 1
        assert checked > 1000

    def test_the_parent_slot_is_invisible_to_the_serializer(
        self, parser, serializer, document_bytes, context
    ):
        first = parser.from_bytes(document_bytes, wml.Document)
        second = parser.from_bytes(document_bytes, wml.Document)
        link_parents(second, context=context)
        assert serializer.render(first) == serializer.render(second)

    def test_equality_ignores_the_parent_on_generated_classes(
        self, parser, document_bytes, context
    ):
        a = parser.from_bytes(document_bytes, wml.Document)
        b = parser.from_bytes(document_bytes, wml.Document)
        link_parents(b, context=context)
        assert a == b

    def test_deep_copy_of_a_parsed_subtree(self, document, context, serializer):
        link_parents(document, context=context)
        paragraph = next(
            node for node in iter_tree(document, context=context)
            if isinstance(node, wml.P)
        )
        assert paragraph.get_parent() is not None
        c = deep_copy(paragraph)
        assert c == paragraph
        assert c.get_parent() is None
        assert serializer.render(c) == serializer.render(paragraph)
        for node in iter_tree(c, context=context):
            for f in dataclasses.fields(node):
                value = getattr(node, f.name, None)
                items = value if isinstance(value, (list, tuple)) else (value,)
                for item in items:
                    if isinstance(item, Child):
                        assert item.get_parent() is node

    def test_mc_ignorable_survives_on_a_font_table(self, parser, serializer):
        """The anyAttribute schema patch, from the object model's side."""
        with zipfile.ZipFile(ROOT / "samples" / "tables.docx") as z:
            data = z.read("word/fontTable.xml")
        fonts = parser.from_bytes(data, wml.Fonts)
        assert fonts.other_attributes
        assert any("Ignorable" in k for k in fonts.other_attributes)
        assert "Ignorable" in serializer.render(fonts)
