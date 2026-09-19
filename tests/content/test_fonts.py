"""``RunFontSelector``, ``Mapper`` and the font tables. CR-002 sections 6.3 and 15.3.

The parity with docx4j is ``test_fonts_goldens.py``; this is what the goldens
cannot say: the character-range table branch by branch, the rules docx4j's
CR-016 measured against Word one at a time, the mapper's precedence, the tables'
own shape, and the wiring into the package.
"""

from __future__ import annotations

import pytest
from conftest import sample

from docx4j_py import create_package, load
from docx4j_py.model.fonts import (
    DEFAULT_FONT_REGISTRY,
    THEME_2007,
    THEME_2013,
    THEME_2023,
    IdentityPlusMapper,
    PhysicalFont,
    RunFontSelector,
    SimpleFontRegistry,
    class_of,
    get_script_for_language_tag,
    has_bold_face,
    is_known_family,
    script_of,
    span_script,
    substitution_class,
    symbol_font_name,
    to_english,
    word_default_for,
)
from docx4j_py.model.fonts.data import (
    EAST_ASIAN_FAMILIES,
    LINE_METRIC_FAMILIES,
    MS_FONTS_WITH_BOLD,
    MS_FONT_NAMES,
    SUBSTITUTE_ROWS,
    VCL_ENTRIES,
)
from docx4j_py.wml import wml

# ---------------------------------------------------------------------------
# the tables
# ---------------------------------------------------------------------------


def test_the_tables_are_docx4js():
    """The counts ``scripts/generate_fonts.py`` printed, pinned."""
    assert len(SUBSTITUTE_ROWS) == 30
    assert len(VCL_ENTRIES) == 419
    # 137 <font> entries, of which Marlett is listed twice
    assert len(MS_FONT_NAMES()) == 136
    assert len(MS_FONTS_WITH_BOLD()) == 41
    assert len(LINE_METRIC_FAMILIES()) == 512
    assert len(EAST_ASIAN_FAMILIES()) == 46
    assert len(DEFAULT_FONT_REGISTRY.all()) == 43


def test_the_default_registry_is_the_font_jars():
    assert DEFAULT_FONT_REGISTRY.get("Carlito Regular").family_name == "Carlito"
    # case-insensitive, and the twin suffixes are stripped
    assert DEFAULT_FONT_REGISTRY.get("carlito regular+kerned+nobold") is not None
    assert DEFAULT_FONT_REGISTRY.get("Calibri") is None, "the jars carry no Calibri"
    assert DEFAULT_FONT_REGISTRY.get(None) is None


def test_the_classes_and_the_known_families():
    assert class_of("Georgia") == "SERIF"
    assert class_of("Segoe UI") == "SANS"
    assert class_of("Consolas") == "MONO"
    # "known" is substitution_class, not class_of: a name that merely ends in
    # Sans is not a known family, so the Word-default pass acts on it
    assert substitution_class("Acme Sans") == "UNKNOWN"
    assert class_of("Acme Sans") == "SANS"
    assert is_known_family("Acme Sans") is False
    assert is_known_family("Arial") is True
    # a family left to the document default is "known" for this purpose
    assert is_known_family("Arial Narrow") is True


def test_has_bold_face():
    assert has_bold_face("Arial") is True
    assert has_bold_face("Calibri Light") is False, "a weight word ends the name"
    assert has_bold_face("Segoe UI Semibold") is False
    assert has_bold_face("Some Corporate Face") is True


# ---------------------------------------------------------------------------
# the character-range table (MS-OI29500 17.3.2.26)
# ---------------------------------------------------------------------------


@pytest.fixture
def selector():
    return RunFontSelector()


@pytest.mark.parametrize(
    ("code_point", "expected"),
    [
        (0x0020, "A"),  # U+0020 is ASCII: the space between two CJK words is Latin
        (0x0041, "A"),
        (0x00E9, "H"),  # Latin-1 Supplement: hAnsi
        (0x0100, "H"),  # Latin Extended-A
        (0x0400, "H"),  # Cyrillic
        (0x05D0, "A"),  # Hebrew: ascii, measured (fonts-hebrew-no-cs)
        (0x0627, "A"),  # Arabic
        (0x0905, "C"),  # Devanagari: cs (docx4j issues 622, 666)
        (0x0E01, "C"),  # Thai
        (0x1000, "C"),  # Myanmar
        (0x1780, "C"),  # Khmer
        (0x1100, "E"),  # Hangul Jamo
        (0x1E00, "H"),  # Latin Extended Additional
        (0x2018, "H"),  # General Punctuation
        (0x2192, "H"),  # Arrows: hAnsi, the glyph check not being ported
        (0x3042, "E"),  # Hiragana
        (0x4E2D, "E"),  # CJK
        (0xF0FC, "H"),  # Private Use Area
        (0xF900, "E"),  # CJK Compatibility Ideographs
        (0xFB00, "H"),  # Alphabetic Presentation Forms, below FB1D
        (0xFB1D, "A"),  # ...and above it
        (0xFB50, "A"),  # Arabic Presentation Forms-A
        (0xFE30, "E"),  # CJK Compatibility Forms
        (0xFE70, "A"),  # Arabic Presentation Forms-B
        (0xFF0C, "E"),  # Halfwidth and Fullwidth Forms
        (0x10A0, "H"),  # Georgian: "all ranges not listed" is hAnsi
        (0x1F600, "H"),  # astral: hAnsi where no emoji font is configured
    ],
)
def test_the_range_table(selector, code_point, expected):
    answer = selector.font_for(code_point, None, None, "E", "A", "H", "C")
    assert answer == expected


def test_the_hint_exceptions_of_the_latin_1_supplement(selector):
    # "If hint is eastAsia, the following characters use eastAsia: A1, A4, ..."
    assert selector.font_for(0xA1, "eastAsia", None, "E", "A", "H", "C") == "E"
    assert selector.font_for(0xA2, "eastAsia", None, "E", "A", "H", "C") == "H"
    # ...and with Chinese, E0-E1, E8-EA, EC-ED, F2-F3, F9-FA, FC as well
    assert selector.font_for(0xE0, "eastAsia", "zh-CN", "E", "A", "H", "C") == "E"
    assert selector.font_for(0xE0, "eastAsia", "en-US", "E", "A", "H", "C") == "H"
    # a run naming no eastAsia font gets hAnsi, never nothing
    assert selector.font_for(0x4E2D, None, None, None, "A", "H", "C") == "H"


def test_the_emoji_font_is_an_option_not_a_property():
    plain = RunFontSelector()
    assert plain.font_for(0x1F600, None, None, "E", "A", "H", "C") == "H"
    with_emoji = RunFontSelector(emoji_font="Segoe UI Emoji")
    assert with_emoji.font_for(0x1F600, None, None, "E", "A", "H", "C") == "Segoe UI Emoji"


# ---------------------------------------------------------------------------
# the run's dispatch
# ---------------------------------------------------------------------------


def _spans(selector, xml, text):
    r_pr = wml(xml)
    return selector.spans(None, r_pr, text, r_pr_is_effective=True)


def test_cs_by_value_and_rtl_alone(selector):
    """docx4j 17.1.1: ``w:cs w:val="0"`` turns an inherited flag off."""
    on = _spans(selector, "<w:rPr><w:cs/><w:rFonts w:cs='Arial'/></w:rPr>", "abc")
    assert on[0].document_font == "Arial"
    off = _spans(selector, "<w:rPr><w:cs w:val='0'/><w:rFonts w:cs='Arial'/></w:rPr>", "abc")
    assert off[0].document_font != "Arial", "an explicit false is off"
    # w:rtl alone on Latin text still takes the cs font
    rtl = _spans(selector, "<w:rPr><w:rtl/><w:rFonts w:cs='Arial'/></w:rPr>", "abc")
    assert rtl[0].document_font == "Arial"


def test_a_symbol_font_takes_the_whole_run_whatever_case_it_is_written_in(selector):
    assert symbol_font_name("symbol") == "Symbol"
    assert symbol_font_name("WINGDINGS 2") == "Wingdings 2"
    assert symbol_font_name("Arial") is None
    spans = _spans(selector, "<w:rPr><w:rFonts w:hAnsi='wingdings'/></w:rPr>", "abc")
    assert [s.document_font for s in spans] == ["Wingdings"]


def test_the_preamble_rule(selector):
    xml = "<w:rPr><w:rFonts w:ascii='X' w:hAnsi='X' w:eastAsia='Times New Roman'/></w:rPr>"
    spans = _spans(selector, xml, "a中b")
    assert [s.document_font for s in spans] == ["X"], "one span, the ascii font"


def test_a_run_with_no_rfonts_gets_the_default_in_the_two_latin_slots(selector):
    """Not the eastAsia slot: a Times New Roman there would fire the preamble rule."""
    spans = selector.spans(None, wml("<w:rPr/>"), "a中b", r_pr_is_effective=True)
    # the default font is Times New Roman with no resolver, and the CJK character
    # falls to the (unnamed) eastAsia slot, which answers hAnsi
    assert [s.text for s in spans] == ["a中b"]
    assert spans[0].document_font == "Times New Roman"


def test_spans_fold_by_font_and_report_the_script(selector):
    xml = "<w:rPr><w:rFonts w:ascii='A' w:hAnsi='H' w:eastAsia='E'/></w:rPr>"
    spans = _spans(selector, xml, "ab中中cd")
    assert [(s.text, s.document_font) for s in spans] == [
        ("ab", "A"),
        ("中中", "E"),
        ("cd", "A"),
    ]
    assert [s.script for s in spans] == ["LATIN", "CJK", "LATIN"]


def test_span_script_groups_the_cjk_scripts_and_shares_the_rest():
    assert span_script(ord("a")) == "LATIN"
    assert span_script(ord(" ")) is None, "a shared character joins what it follows"
    assert span_script(ord("中")) == "CJK"
    assert span_script(ord("あ")) == "CJK"
    assert span_script(ord("가")) == "CJK"
    assert span_script(0x2192) == "SYMBOL"
    assert span_script(0x1F600) == "EMOJI"
    assert script_of(ord("א")) == "HEBREW"
    assert script_of(ord("ا")) == "ARABIC"


# ---------------------------------------------------------------------------
# themes
# ---------------------------------------------------------------------------


def test_the_theme_language_is_matched_by_exact_subtag():
    """docx4j CR-016 probe ``fonts-theme-lang``: Estonian is not Ethiopic."""
    assert get_script_for_language_tag("et-EE") is None
    assert get_script_for_language_tag("ti-ET") == "Ethi"
    assert get_script_for_language_tag("ja-JP") == "Jpan"
    assert get_script_for_language_tag("zh-CN") == "Hans"
    assert get_script_for_language_tag("zh-TW") == "Hant"
    assert get_script_for_language_tag(None) is None


def test_a_themeless_package_answers_the_office_theme_for_the_two_latin_slots():
    for theme, minor, major in (
        (THEME_2023, "Aptos", "Aptos Display"),
        (THEME_2013, "Calibri", "Calibri Light"),
        (THEME_2007, "Calibri", "Cambria"),
    ):
        selector = RunFontSelector(default_theme=theme.value)
        assert selector.theme_font("minorHAnsi") == minor
        assert selector.theme_font("majorHAnsi") == major
        # the East Asian and complex-script references resolve to nothing, so the
        # explicit attribute beside them stands
        assert selector.theme_font("minorEastAsia") is None
        assert selector.theme_font("minorBidi") is None


def test_cjk_names_are_collected_in_english():
    assert to_english("宋体") == "SimSun"
    assert to_english("Arial") is None


# ---------------------------------------------------------------------------
# the mapper
# ---------------------------------------------------------------------------


def test_the_precedence_installed_then_the_mappers_own_then_the_clones():
    mapper = IdentityPlusMapper()
    mapper.populate(["Carlito", "Calibri", "Malgun Gothic"], None)
    assert mapper.get_decision("Carlito").source == "MAPPER_OWN"
    assert mapper.get_decision("Calibri").source == "METRIC_CLONE"
    assert mapper.get_decision("Calibri").physical_font.name == "Carlito Regular"
    assert mapper.get_decision("Malgun Gothic").source == "UNMAPPED"
    assert mapper.get("Malgun Gothic") is None


def test_an_installed_font_beats_an_embedded_one():
    mapper = IdentityPlusMapper()
    embedded = PhysicalFont("Carlito Regular (embedded)")
    mapper.register_regular_form("Carlito Regular", embedded)
    mapper.populate(["Carlito Regular"], None)
    assert mapper.get("Carlito Regular").name == "Carlito Regular"
    assert mapper.get_decision("Carlito Regular").source == "INSTALLED"


def test_the_alt_name_chain_and_its_cycle_guard():
    fonts = wml(
        "<w:fonts>"
        "<w:font w:name='ProbeD'><w:altName w:val='ProbeE'/></w:font>"
        "<w:font w:name='ProbeE'><w:altName w:val='Arial'/></w:font>"
        "<w:font w:name='Loop1'><w:altName w:val='Loop2'/></w:font>"
        "<w:font w:name='Loop2'><w:altName w:val='Loop1'/></w:font>"
        "</w:fonts>"
    )
    mapper = IdentityPlusMapper()
    mapper.populate(["ProbeD", "Loop1"], fonts)
    decision = mapper.get_decision("ProbeD")
    assert decision.source == "ALT_NAME"
    assert decision.via == "w:altName ProbeE -> Arial"
    assert decision.physical_font.name == "Arimo Regular"
    # a cycle stops rather than spinning
    assert mapper.get_decision("Loop1").source in ("UNMAPPED", "WORD_DEFAULT")


def test_word_default_for_reads_the_font_tables_family():
    table = {
        "probea": wml("<w:font w:name='ProbeA'><w:family w:val='roman'/></w:font>"),
        "probeb": wml("<w:font w:name='ProbeB'><w:family w:val='swiss'/></w:font>"),
        "probec": wml("<w:font w:name='ProbeC'/>"),
    }
    assert word_default_for("ProbeA", table) == "Cambria"
    assert word_default_for("ProbeB", table) == "Calibri"
    assert word_default_for("ProbeC", table) == "Calibri"
    assert word_default_for("Nowhere", table) == "Cambria", "no entry at all"


def test_the_no_bold_face_alias_and_the_word_defaulted_exception():
    mapper = IdentityPlusMapper()
    mapper.populate(["Calibri Light"], None)
    font = mapper.get("Calibri Light")
    assert font.no_bold_face is True
    assert font.name.endswith("+nobold")
    assert font.family_name == "Carlito"

    # a font Word itself could not find is substituted whole, its real bold included
    fonts = wml("<w:fonts><w:font w:name='EnBW DIN Pro Light'/></w:fonts>")
    other = IdentityPlusMapper()
    other.populate(["EnBW DIN Pro Light"], fonts)
    assert other.get_decision("EnBW DIN Pro Light").source == "WORD_DEFAULT"
    assert other.get("EnBW DIN Pro Light").no_bold_face is False


def test_a_registry_with_nothing_in_it_maps_nothing():
    mapper = IdentityPlusMapper(SimpleFontRegistry())
    mapper.populate(["Calibri", "Arial"], None)
    assert [d.source for d in mapper.get_decisions()] == ["UNMAPPED", "UNMAPPED"]


# ---------------------------------------------------------------------------
# the wiring
# ---------------------------------------------------------------------------


def test_building_the_selector_and_the_mapper_unmarshals_no_side_part():
    package = sample("2010-sample1.docx")
    theme = package.theme_part
    font_table = package.font_table_part
    settings = package.document_settings_part
    before = {}
    for part in (theme, font_table, settings, package.style_definitions_part):
        if part is not None:
            before[str(part.part_name)] = part.xml

    package.get_font_mapper()

    for part in (theme, font_table, settings, package.style_definitions_part):
        if part is None:
            continue
        assert part.is_unmarshalled is False, f"{part.part_name} was unmarshalled"
        assert part.xml == before[str(part.part_name)]


def test_the_selector_and_the_mapper_are_one_per_package():
    package = create_package()
    assert package.get_run_font_selector() is package.get_run_font_selector()
    assert package.get_font_mapper() is package.font_mapper
    # a created package has a theme part, so its default font is the theme's
    assert package.get_run_font_selector().default_font == "Aptos"


def test_font_name_reports_the_resolved_face():
    package = create_package()
    paragraph = package.body.insert_paragraph("x")
    assert paragraph.font.name == "Aptos", "from the theme, through w:asciiTheme"
    assert paragraph.get_font(direct=True).name == "", "the run says nothing"


def test_fonts_in_use_walks_every_story():
    package = load("tests/fixtures/numbering-stories.docx")
    fonts = package.main_document_part.fonts_in_use()
    assert fonts, "a document uses at least its default font"
    assert all(isinstance(name, str) and name.strip() for name in fonts)
    styles = package.main_document_part.styles_in_use()
    assert "Normal" in styles or styles
