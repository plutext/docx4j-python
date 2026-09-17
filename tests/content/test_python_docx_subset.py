"""The python-docx compatibility promise, kept honest.

CR-003 section 3.6 and section 20: ``tests/python_docx_subset.json`` is the
committed list of the python-docx members :mod:`docx4j_py.docx` implements,
derived by ``scripts/python_docx_subset.py`` from an **installed** python-docx
at the version the list pins. This test asserts that every member the list
marks implemented is really on the facade class, with the right kind
(``property`` or ``method``), and prints the ones deliberately omitted with
their reasons.

Built the same way as ``test_office_js_subset.py``, and for the same reason:
the promise lives with the implementation, and python-docx's **own** test suite
is not run (decided question 7 --- its tests exercise its lxml internals, which
were never the promise).
"""

from __future__ import annotations

import json

import pytest
from conftest import ROOT

from docx4j_py import docx as facade
from docx4j_py.docx import (
    Cm,
    Emu,
    Inches,
    Length,
    Mm,
    Pt,
    RGBColor,
    Twips,
)

SUBSET = ROOT / "tests" / "python_docx_subset.json"

#: The facade class each python-docx interface is. Every name in the committed
#: list must be here, which the first test asserts.
CLASSES: dict[str, type] = {
    "Document": facade.Document,
    "Paragraph": facade.Paragraph,
    "Run": facade.Run,
    "Font": facade.Font,
    "ParagraphFormat": facade.ParagraphFormat,
    "TabStops": facade.TabStops,
    "TabStop": facade.TabStop,
    "ColorFormat": facade.ColorFormat,
    "Hyperlink": facade.Hyperlink,
    "RenderedPageBreak": facade.RenderedPageBreak,
    "Table": facade.Table,
    "_Row": facade._Row,
    "_Rows": facade._Rows,
    "_Column": facade._Column,
    "_Columns": facade._Columns,
    "_Cell": facade._Cell,
    "Section": facade.Section,
    "Sections": facade.Sections,
    "_Header": facade._Header,
    "_Footer": facade._Footer,
    "InlineShape": facade.InlineShape,
    "InlineShapes": facade.InlineShapes,
    "Styles": facade.Styles,
    "BaseStyle": facade.BaseStyle,
    "CharacterStyle": facade.CharacterStyle,
    "ParagraphStyle": facade.ParagraphStyle,
    "_TableStyle": facade._TableStyle,
    "_NumberingStyle": facade._NumberingStyle,
    "CoreProperties": facade.CoreProperties,
    "Settings": facade.Settings,
    "Comment": facade.Comment,
    "Comments": facade.Comments,
    "Length": Length,
    "RGBColor": RGBColor,
    "Drawing": facade.Drawing,
}

#: The ``Length`` constructors, by the name the list uses.
CONSTRUCTORS: dict[str, type] = {
    "Emu": Emu,
    "Twips": Twips,
    "Pt": Pt,
    "Mm": Mm,
    "Cm": Cm,
    "Inches": Inches,
}


@pytest.fixture(scope="module")
def subset() -> dict:
    """The committed list."""
    return json.loads(SUBSET.read_text(encoding="utf-8"))


def kind_of(cls: type, name: str) -> str | None:
    """``"property"``, ``"method"`` or None, by what the class declares.

    A ``__slots__`` entry counts as a property, as in
    ``test_office_js_subset.py``: what is being asserted is the shape of the
    member, not how the class happens to store it.
    """
    for klass in cls.__mro__:
        if name in vars(klass):
            member = vars(klass)[name]
            is_data = isinstance(member, property) or type(member).__name__ == "member_descriptor"
            return "property" if is_data else "method"
    return None


def test_the_list_is_committed_and_names_its_source(subset):
    assert SUBSET.exists()
    assert subset["source"] == "python-docx 1.2.0, introspected"
    assert subset["version"] == "1.2.0"
    assert set(subset["interfaces"]) == set(CLASSES)


def test_every_implemented_member_is_there_with_the_right_kind(subset):
    missing: list[str] = []
    wrong_kind: list[str] = []
    omitted: list[str] = []

    for interface, members in subset["interfaces"].items():
        cls = CLASSES[interface]
        for member in members:
            where = f"{interface}.{member['name']}"
            if not member["implemented"]:
                omitted.append(where)
                assert member.get("reason"), f"{where} is omitted with no reason"
                continue
            kind = kind_of(cls, member["name"])
            if kind is None:
                missing.append(where)
            elif kind != member["kind"]:
                wrong_kind.append(
                    f"{where} is a {kind}, python-docx has a {member['kind']}"
                )

    print(f"\ndeliberately omitted ({len(omitted)}): {', '.join(sorted(omitted))}")
    assert not missing and not wrong_kind, (
        f"the facade owes {missing or 'nothing'}; wrong kind: {wrong_kind or 'none'}"
    )


def test_the_omissions_are_exactly_the_ones_recorded(subset):
    """The omissions are a closed list, so adding one is a deliberate act."""
    omitted = {
        f"{interface}.{member['name']}"
        for interface, members in subset["interfaces"].items()
        for member in members
        if not member["implemented"]
    }
    assert omitted == {
        "Comments.add_comment",
        "Drawing.image",
        "RenderedPageBreak.following_paragraph_fragment",
        "RenderedPageBreak.preceding_paragraph_fragment",
        "Run.mark_comment_range",
        "Styles.latent_styles",
        "_Cell.merge",
    }


def test_an_omitted_member_raises_rather_than_answering_wrongly():
    """What the facade does not offer says so, or is simply not there."""
    from docx4j_py.docx import Document

    document = Document()
    with pytest.raises(AttributeError, match="latent_styles"):
        _ = document.styles.latent_styles
    # the rest are not there at all, which is what "no proxies" means
    paragraph = document.add_paragraph("x")
    assert hasattr(paragraph.paragraph_format, "tab_stops"), "offered since 2026-09-18"
    assert not hasattr(paragraph.add_run("y"), "mark_comment_range")


def test_the_cell_merge_refusal_says_what_to_do_instead(new_package):
    """``_Cell.merge`` raises ``NotImplementedError`` with the way round it."""
    from docx4j_py.docx import Document

    document = Document()
    table = document.add_table(2, 2)
    with pytest.raises(NotImplementedError, match="grid_span"):
        table.cell(0, 0).merge(table.cell(1, 1))


def test_the_length_constructors_are_python_docxs_to_the_emu(subset):
    """``Pt(24)`` is 304800 EMU, and every unit's factor is the list's."""
    for name, factor in subset["constructors"].items():
        constructor = CONSTRUCTORS[name]
        assert issubclass(constructor, Length), name
        assert constructor(1) == factor, name
        assert constructor(3) == 3 * factor, name
    assert Pt(24) == 304800
    assert Inches(1) == 914400
    assert Length(304800).pt == 24.0
    assert Length(914400).inches == 1.0
    assert Length(360000).cm == 1.0
    assert Length(36000).mm == 1.0
    assert Length(635).twips == 1
    assert Length(914400).emu == 914400


def test_the_enums_have_python_docxs_members_and_w_val_strings(subset):
    """Every member name, its MS API integer and its ``w:val``, all three."""
    for name, declared in subset["enums"].items():
        ours = getattr(facade, name)
        for member_name, expected in declared["members"].items():
            member = getattr(ours, member_name, None)
            assert member is not None, f"{name}.{member_name}"
            assert member.value == expected["value"], f"{name}.{member_name} value"
            assert getattr(member, "xml_value", None) == expected["xml_value"], (
                f"{name}.{member_name} xml_value"
            )
        assert {m.name for m in ours} == set(declared["members"]), name


def test_the_aliases_python_docx_code_imports_are_all_there():
    """``WD_ALIGN_PARAGRAPH`` and the rest are the same objects as their long names."""
    assert facade.WD_ALIGN_PARAGRAPH is facade.WD_PARAGRAPH_ALIGNMENT
    assert facade.WD_BREAK is facade.WD_BREAK_TYPE
    assert facade.WD_COLOR is facade.WD_COLOR_INDEX
    assert facade.WD_ORIENT is facade.WD_ORIENTATION
    assert facade.WD_SECTION is facade.WD_SECTION_START
    assert facade.WD_ALIGN_VERTICAL is facade.WD_CELL_VERTICAL_ALIGNMENT
    assert facade.WD_ROW_HEIGHT is facade.WD_ROW_HEIGHT_RULE
    assert facade.WD_HEADER_FOOTER is facade.WD_HEADER_FOOTER_INDEX
    assert facade.WD_INLINE_SHAPE is facade.WD_INLINE_SHAPE_TYPE
    assert facade.MSO_THEME_COLOR is facade.MSO_THEME_COLOR_INDEX


def test_the_omitted_enumeration_is_recorded_with_its_replacement(subset):
    """``WD_BUILTIN_STYLE`` is the one enumeration left out, and says why."""
    assert set(subset["omitted_enums"]) == {"WD_BUILTIN_STYLE"}
    assert "BUILT_IN_STYLES" in subset["omitted_enums"]["WD_BUILTIN_STYLE"]
    assert not hasattr(facade, "WD_BUILTIN_STYLE")
    assert not hasattr(facade, "WD_STYLE")


def test_the_script_would_write_what_is_committed():
    """Run the generator when python-docx is importable; skip when it is not."""
    import sys

    pytest.importorskip("docx", reason="python-docx is not installed here, and is not a dependency")
    sys.path.insert(0, str(ROOT / "scripts"))
    import python_docx_subset

    assert python_docx_subset.main(["--check"]) == 0


def test_the_committed_file_is_valid_json_the_repository_tracks(subset):
    assert isinstance(subset, dict)
    counts = subset["counts"]
    assert counts["interfaces"] == len(subset["interfaces"])
    assert counts["members"] == sum(len(m) for m in subset["interfaces"].values())
    assert counts["implemented"] == counts["members"] - 7
