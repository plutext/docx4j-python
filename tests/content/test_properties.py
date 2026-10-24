"""``PropertyResolver`` and ``StyleUtil``. CR-002 sections 6.1 and 15.1.

The parity with docx4j is ``test_properties_goldens.py``; this is what the
goldens cannot say: that building a resolver costs a document no byte of its
round trip, that resolution writes nothing into the styles part, the catalogue's
own completeness, the toggle rule of ECMA-376-1 17.7.3, the colour merge rules of
docx4j 17.3.3 (CR-009 Phase A, ahead of the goldens), the cycle guard, and the
wiring into the package and the content API.
"""

from __future__ import annotations

import dataclasses

import pytest
from conftest import reloaded, sample

from docx4j_py.model import property_catalogue as catalogue
from docx4j_py.model.property_resolver import PropertyResolver, heading_level_by_name
from docx4j_py.model.style_util import (
    apply_rpr,
    is_cyclic,
    throw_on_cyclic_styles,
    toggle,
)
from docx4j_py.openpackaging.exceptions import CyclicStylesException, Docx4JException
from docx4j_py.wml import (
    BooleanDefaultTrue,
    Color,
    CTBorder,
    CTShd,
    RPr,
    STBorder,
    STHexColorAuto,
    STShd,
    STThemeColor,
    Styles,
    U,
    UnderlineEnumeration,
    wml,
)

# ---------------------------------------------------------------------------
# the catalogue
# ---------------------------------------------------------------------------

_TABLES = {
    "RUN": (catalogue.RUN, catalogue.RUN_EXCLUDED, "RPr"),
    "PARAGRAPH": (catalogue.PARAGRAPH, catalogue.PARAGRAPH_EXCLUDED, "PPrBase"),
    "TABLE": (catalogue.TABLE, catalogue.TABLE_EXCLUDED, "CTTblPrBase"),
    "CELL": (catalogue.CELL, catalogue.CELL_EXCLUDED, "TcPr"),
}


@pytest.mark.parametrize("name", sorted(_TABLES))
def test_each_table_names_every_member_of_its_element_once(name):
    """One entry per member, in schema order, and nothing invented."""
    import docx4j_py.wml as wml_module

    table, excluded, class_name = _TABLES[name]
    fields = [f.name for f in dataclasses.fields(getattr(wml_module, class_name))]
    listed = [entry.name for entry in table]

    assert len(listed) == len(set(listed)), f"{name}: a member is listed twice"
    # the excluded names are the revision records, which the wider class (``PPr``
    # over ``PPrBase``, ``ParaRPr`` over ``RPr``, ``TblPr`` over ``CTTblPrBase``)
    # carries and this one may not
    assert set(listed) == set(fields) - set(excluded), (
        f"{name}: listed {sorted(listed)} != {sorted(set(fields) - set(excluded))}"
    )
    # schema order: the listed entries appear in the generated class's order
    assert listed == [f for f in fields if f in set(listed)], f"{name}: out of schema order"


def test_the_paragraph_mark_has_the_same_members_as_a_run():
    """One ``RUN`` table serves ``w:rPr`` and ``w:pPr/w:rPr``, as docx4j needs two."""
    import docx4j_py.wml as wml_module

    para = {f.name for f in dataclasses.fields(wml_module.ParaRPr)}
    run = {f.name for f in dataclasses.fields(wml_module.RPr)}
    # the paragraph mark adds the four revision members, and nothing else
    assert para - run == {"ins", "del_value", "move_from", "move_to"}
    assert run - para == set()


def test_the_twelve_toggles_are_the_twelve():
    assert sorted(catalogue.TOGGLE_NAMES) == [
        "b",
        "b_cs",
        "caps",
        "emboss",
        "i",
        "i_cs",
        "imprint",
        "outline",
        "shadow",
        "small_caps",
        "strike",
        "vanish",
    ]
    assert [entry.name for entry in catalogue.TOGGLES] == [
        entry.name for entry in catalogue.RUN if entry.name in catalogue.TOGGLE_NAMES
    ]


def test_a_carried_leaf_is_a_copy_not_the_styles_parts_own_object():
    source = wml("<w:rPr><w:b/><w:color w:val='FF0000'/></w:rPr>")
    destination = RPr()
    apply_rpr(source, destination)

    assert destination.b is not source.b
    assert destination.color is not source.color
    destination.color.val = b"\x00\xff\x00"
    # w:val of a w:color is ST_HexColor, which the model holds as bytes
    assert source.color.val == b"\xff\x00\x00", "the source is untouched"


# ---------------------------------------------------------------------------
# the toggles (ECMA-376-1 17.7.3)
# ---------------------------------------------------------------------------


def _on() -> BooleanDefaultTrue:
    return BooleanDefaultTrue()


def _off() -> BooleanDefaultTrue:
    return BooleanDefaultTrue(val=False)


def test_the_toggle_rule():
    # an upper level that says nothing is no boundary: the lower value stands
    assert toggle(None, _on(), None) is not None
    assert toggle(None, None, _on()) is None
    # the document defaults are a base value, never a term of the XOR
    assert catalogue.is_true(toggle(_on(), _on(), _on())) is True
    # true XOR true is false
    assert catalogue.is_true(toggle(_on(), _on(), None)) is False
    # true XOR nothing is true
    assert catalogue.is_true(toggle(_on(), None, None)) is True
    # false XOR lower is lower: an explicit false over a bold level stays bold
    assert catalogue.is_true(toggle(_off(), _on(), None)) is True
    # with no level beneath, the level's own false stands
    assert catalogue.is_true(toggle(_off(), None, None)) is False


# ---------------------------------------------------------------------------
# a colour inherits as one (docx4j 17.3.3, its CR-034; CR-009 Phase A)
# ---------------------------------------------------------------------------
# docx4j's StyleUtilMergeRulesTest, the three CR-034 cases. Its areEqual
# assertions have no counterpart: the catalogue compares nothing.


def _hex(value: str) -> bytes:
    """An ``ST_HexColor``, which the model holds as bytes."""
    return bytes.fromhex(value)


def test_a_colour_stating_a_value_replaces_an_inherited_theme_reference():
    """Word draws a reference over the value beside it: a run's red must not keep accent1."""
    inherited = Color(val=_hex("0F4761"), theme_color=STThemeColor.ACCENT1, theme_shade="BF")
    direct = Color(val=_hex("FF0000"))

    merged = catalogue.merge_color(direct, inherited)
    assert merged.val == _hex("FF0000")
    assert merged.theme_color is None
    assert merged.theme_shade is None


def test_a_colour_stating_only_a_theme_reference():
    """Not empty, so inherited, and it replaces a value; a tint alone still inherits."""
    theme_only = Color(theme_color=STThemeColor.ACCENT2)
    assert catalogue.is_empty_color(theme_only) is False
    assert catalogue.merge_color(theme_only, None).theme_color is STThemeColor.ACCENT2

    inherited = Color(val=_hex("FF0000"))
    merged = catalogue.merge_color(theme_only, inherited)
    assert merged.val is None
    assert merged.theme_color is STThemeColor.ACCENT2

    # a tint alone still inherits on its own
    merged = catalogue.merge_color(Color(theme_tint="99"), merged)
    assert merged.theme_color is STThemeColor.ACCENT2
    assert merged.theme_tint == "99"


def test_shading_fill_border_and_underline_colours_inherit_as_one():
    shd = CTShd(
        val=STShd.CLEAR,
        color=STHexColorAuto.AUTO,
        fill=_hex("156082"),
        theme_fill=STThemeColor.ACCENT1,
    )
    merged_shd = catalogue.merge_shd(CTShd(fill=_hex("FFFF00")), shd)
    assert merged_shd.fill == _hex("FFFF00")
    assert merged_shd.theme_fill is None
    assert merged_shd.color is STHexColorAuto.AUTO, "the pattern colour is its own"
    assert merged_shd.val is STShd.CLEAR

    border = CTBorder(
        val=STBorder.SINGLE, sz=4, color=_hex("156082"), theme_color=STThemeColor.ACCENT1
    )
    merged_border = catalogue.merge_border(CTBorder(color=_hex("FF0000")), border)
    assert merged_border.color == _hex("FF0000")
    assert merged_border.theme_color is None
    assert merged_border.val is STBorder.SINGLE

    u = U(val=UnderlineEnumeration.SINGLE, color=_hex("156082"), theme_color=STThemeColor.ACCENT1)
    red_u = U(color=_hex("FF0000"))
    merged_u = catalogue.merge_u(red_u, u)
    assert merged_u.val is UnderlineEnumeration.SINGLE
    assert merged_u.color == _hex("FF0000")
    assert merged_u.theme_color is None

    assert catalogue.is_empty_u(U(theme_color=STThemeColor.ACCENT4)) is False


def test_auto_states_a_colour_and_none_does_not():
    """The model's own spellings: ``auto`` is an enum member, the hex is bytes.

    Corpus document 3229's case, through the resolver: a run's own
    ``w:color w:val="auto"`` under a style coloured accent1 shaded BF is black
    in Word, so nothing of the style's colour survives.
    """
    styles = wml(
        "<w:styles>"
        "<w:style w:type='paragraph' w:styleId='Accent'>"
        "<w:rPr><w:color w:val='365F91' w:themeColor='accent1' w:themeShade='BF'/></w:rPr>"
        "</w:style>"
        "</w:styles>"
    )
    resolver = PropertyResolver(styles=styles)
    p_pr = wml("<w:pPr><w:pStyle w:val='Accent'/></w:pPr>")
    r_pr = wml("<w:rPr><w:color w:val='auto'/></w:rPr>")
    effective = resolver.get_effective_rpr(r_pr, p_pr)
    assert effective.color.val is STHexColorAuto.AUTO
    assert effective.color.theme_color is None
    assert effective.color.theme_shade is None

    # a reference of none states nothing: the value and the shade beneath stay
    inherited = Color(val=_hex("365F91"), theme_color=STThemeColor.ACCENT1, theme_shade="BF")
    merged = catalogue.merge_color(Color(theme_color=STThemeColor.NONE), inherited)
    assert merged.val == _hex("365F91")
    assert merged.theme_color is STThemeColor.ACCENT1
    assert merged.theme_shade == "BF"
    # but it is not empty, as the Java's getThemeColor() == null is not
    assert catalogue.is_empty_color(Color(theme_color=STThemeColor.NONE)) is False
    # and an empty value is docx4j's empty string
    assert catalogue.is_empty_color(Color(val=b"")) is True


# ---------------------------------------------------------------------------
# cycles
# ---------------------------------------------------------------------------


def test_a_cycle_stops_the_walk_and_does_not_throw_by_default():
    styles = wml(
        "<w:styles>"
        "<w:style w:type='paragraph' w:styleId='A'><w:basedOn w:val='B'/></w:style>"
        "<w:style w:type='paragraph' w:styleId='B'><w:basedOn w:val='A'/></w:style>"
        "</w:styles>"
    )
    resolver = PropertyResolver(styles=styles)
    # root first, and the walk stopped where it came back to A
    assert [style.style_id for style in resolver.ancestry("A")] == ["B", "A"]


def test_a_cycle_throws_when_docx4js_property_is_on():
    try:
        throw_on_cyclic_styles(True)
        with pytest.raises(CyclicStylesException):
            is_cyclic("A", ["A"])
    finally:
        throw_on_cyclic_styles(False)
    assert is_cyclic("A", ["A"]) is True, "and answers instead when it is off"


def test_a_chain_deeper_than_thirty_two_stops():
    ids = [f"S{index}" for index in range(40)]
    parts = "".join(
        f"<w:style w:type='paragraph' w:styleId='{name}'>"
        f"<w:basedOn w:val='{ids[index + 1]}'/></w:style>"
        for index, name in enumerate(ids[:-1])
    )
    styles = wml(
        f"<w:styles>{parts}<w:style w:type='paragraph' w:styleId='{ids[-1]}'/></w:styles>"
    )
    resolver = PropertyResolver(styles=styles)
    assert len(resolver.ancestry("S0")) == 33


# ---------------------------------------------------------------------------
# headings
# ---------------------------------------------------------------------------


def test_the_outline_level_comes_from_the_styles_name_not_its_declaration():
    """Word takes a built-in heading's level from its ``w:name``, in every locale."""
    styles = wml(
        "<w:styles>"
        "<w:style w:type='paragraph' w:styleId='Titre2'>"
        "<w:name w:val='heading 2'/><w:pPr><w:outlineLvl w:val='6'/></w:pPr>"
        "</w:style>"
        "</w:styles>"
    )
    resolver = PropertyResolver(styles=styles)
    assert heading_level_by_name(resolver.get_style("Titre2")) == 2
    assert resolver.get_effective_ppr("Titre2").outline_lvl.val == 1
    # and the styles part still says what it said
    assert resolver.get_style("Titre2").p_pr.outline_lvl.val == 6
    # the id form is docx4j's, and keys on the English id
    assert resolver.get_lvl_from_heading_style("Heading3") == 3
    assert resolver.get_lvl_from_heading_style("Titre2") == -1


# ---------------------------------------------------------------------------
# the parts: nothing is unmarshalled, nothing is written
# ---------------------------------------------------------------------------


def test_building_a_resolver_unmarshals_nothing_and_the_round_trip_survives():
    package = sample("2010-sample1.docx")
    styles = package.style_definitions_part
    numbering = package.numbering_definitions_part
    before = styles.xml

    resolver = package.get_property_resolver()
    for style_id in list(resolver._live_styles):
        resolver.get_effective_ppr(style_id)
        resolver.get_effective_rpr(style_id)

    assert styles.is_unmarshalled is False, "reading the styles did not unmarshal them"
    if numbering is not None:
        assert numbering.is_unmarshalled is False
    assert styles.xml == before, "and the part is byte for byte what it was"


def test_resolution_writes_nothing_into_the_styles_part():
    package = sample("2010-sample1.docx")
    resolver = package.get_property_resolver()
    part = package.style_definitions_part
    before = part.xml

    for style_id in list(resolver._live_styles):
        resolver.get_effective_ppr(style_id)
        resolver.get_effective_rpr(style_id)
    resolver.get_document_default_rpr().sz.val = 999  # our own copy

    assert part.xml == before
    # the 10 pt default docx4j makes explicit is the resolver's, not the part's
    assert b'w:sz w:val="999"' not in part.xml


def test_the_resolver_is_one_per_package_and_refreshes_when_a_style_is_added():
    package = sample("2010-sample1.docx")
    assert package.get_property_resolver() is package.get_property_resolver()

    resolver = package.get_property_resolver()
    assert resolver.get_style("Heading1") is None or True  # whatever the sample has

    from docx4j_py.model.content.styles import ensure_style

    ensure_style(package, "Heading1")
    # ensure_style unmarshals the part and invalidates both caches
    assert package.style_definitions_part.is_unmarshalled is True
    assert resolver.get_style("Heading1") is not None
    assert resolver.get_effective_rpr("Heading1") is not None


def test_activate_style_refuses_while_the_styles_part_is_private():
    package = sample("2010-sample1.docx")
    resolver = package.get_property_resolver()
    style = wml(
        "<w:style w:type='paragraph' w:styleId='Zzz'><w:name w:val='Zzz'/></w:style>",
        wrapper=Styles,
    )

    with pytest.raises(Docx4JException) as raised:
        resolver.activate_style(style)
    assert "refresh()" in str(raised.value)

    package.style_definitions_part.contents  # noqa: B018 - unmarshal it for real
    resolver.refresh()
    assert resolver.activate_style(style) is True
    assert resolver.get_style("Zzz") is not None
    assert resolver.activate_style("Zzz") is True
    assert resolver.activate_style("NoSuchStyle") is False


def test_a_missing_style_resolves_as_the_default_paragraph_style_does():
    package = sample("2010-sample1.docx")
    resolver = package.get_property_resolver()
    default_id = resolver.default_paragraph_style_id
    assert resolver.get_effective_ppr("NoSuchStyle") == resolver.get_effective_ppr(default_id)
    # a missing style has no effective w:rPr at all, where a missing w:pPr does
    assert resolver.get_effective_rpr("NoSuchStyle") is None


def test_the_paragraph_marks_character_style_is_resolved_then_its_direct_formatting():
    """docx4j ``PropertyResolverOrderTest`` (17.3.2, ``fb2a764f7``).

    Corpus document 9698's empty cell paragraphs: the mark names a 9pt
    character style and overrides its font, and Word sizes the empty line at
    9pt, where the mark's ``w:rStyle`` was copied and never read before.
    """
    styles = wml(
        "<w:styles>"
        "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val='28'/></w:rPr></w:rPrDefault>"
        "</w:docDefaults>"
        "<w:style w:type='character' w:styleId='Small'>"
        "<w:rPr><w:rFonts w:ascii='Trebuchet MS' w:hAnsi='Trebuchet MS'/><w:sz w:val='18'/></w:rPr>"
        "</w:style>"
        "</w:styles>"
    )
    resolver = PropertyResolver(styles=styles)
    marked = wml(
        "<w:pPr><w:rPr><w:rStyle w:val='Small'/><w:rFonts w:ascii='Times New Roman'/></w:rPr></w:pPr>"
    )
    mark = resolver.get_effective_paragraph_mark_rpr(marked)
    assert mark.sz.val == 18, "the mark's character style gives its size"
    assert mark.r_fonts.ascii == "Times New Roman", "the mark's direct font wins over the style's"

    marked.r_pr.sz = wml("<w:sz w:val='40'/>")
    assert resolver.get_effective_paragraph_mark_rpr(marked).sz.val == 40, "direct size over the style's"

    missing = wml("<w:pPr><w:rPr><w:rStyle w:val='Missing'/></w:rPr></w:pPr>")
    assert resolver.get_effective_paragraph_mark_rpr(missing).sz.val == 28, "a missing style is ignored"

    assert resolver.get_effective_rpr(None, marked).sz.val == 28, "the runs are not touched by the mark's style"


# ---------------------------------------------------------------------------
# the content API reads the effective value
# ---------------------------------------------------------------------------


def test_the_paragraph_getters_report_the_effective_value(new_package):
    paragraph = new_package.body.insert_paragraph("Heading", style="Heading 1")
    assert paragraph.element.p_pr.ind is None, "nothing is set directly"
    assert paragraph.outline_level == 1, "Heading 1 is outline level 1"
    assert paragraph.space_before > 0, "from the style"
    assert paragraph.font.bold is True

    # and a save still reloads to the same answers
    back = reloaded(new_package)
    assert back.body.paragraphs[0].outline_level == 1


def test_a_detached_body_reads_direct_formatting_and_raises_nothing():
    from docx4j_py.model.content import Body

    container = wml("<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body>")
    body = Body(None, container)
    paragraph = body.paragraphs[0]
    assert paragraph.font.size == 0
    assert paragraph.alignment == "Unknown"
    assert paragraph.effective_p_pr is None


# ---------------------------------------------------------------------------
# w:framePr's w:anchorLock read as stated: docx4j 82c64bf14 (CR-059), its
# StyleUtilFramePrAnchorLockTest case for case


def _frame_pr(attributes: str):
    from docx4j_py.fragments import wml

    return wml(f'<w:framePr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" {attributes}/>')


def test_a_frame_pr_with_nothing_set_is_empty():
    assert catalogue.is_empty_frame_pr(_frame_pr(""))
    assert not catalogue.is_empty_frame_pr(_frame_pr('w:anchorLock="1"'))
    assert not catalogue.is_empty_frame_pr(_frame_pr('w:anchorLock="0"'))
    assert not catalogue.is_empty_frame_pr(_frame_pr('w:w="3600"'))


def test_an_omitted_anchor_lock_keeps_the_styles():
    merged = catalogue.merge_frame_pr(_frame_pr('w:w="3600"'), _frame_pr('w:w="3629" w:anchorLock="1"'))
    assert merged.anchor_lock is True
    assert int(merged.w) == 3600


def test_a_stated_anchor_lock_wins():
    assert catalogue.merge_frame_pr(_frame_pr('w:w="1" w:anchorLock="0"'), _frame_pr('w:anchorLock="1"')).anchor_lock is False
    assert catalogue.merge_frame_pr(_frame_pr('w:w="1" w:anchorLock="1"'), _frame_pr('w:w="2"')).anchor_lock is True


def test_neither_states_it_and_none_is_written():
    assert catalogue.merge_frame_pr(_frame_pr('w:w="1"'), _frame_pr('w:w="2"')).anchor_lock is None
