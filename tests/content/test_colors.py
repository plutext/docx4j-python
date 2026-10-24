"""Theme colours as Word draws them: CR-009 Phase B, docx4j's ``org.docx4j.model.color``, the WordprocessingML half.

The vectors are docx4j's (``tests/fixtures/theme-colors/``, copied from
``docx4j-core-tests/src/test/resources/theme-colors/`` at ``a9e0303ff`` with a
``SHA1SUMS``, each file the same blob docx4j added in ``3dbe4eee5`` and the same
SHA-1 as docx4j-core-ts's ``test/theme-colors/manifest.json`` at ``76fb45d``):
the P1 and P2 probes, what Word 365 drew for each case (``readings.tsv``, read
from its PDF), the cases (``cases.tsv``) and the 117 distinct tint and shade
triples of four real-document corpora (``word-tint-shade.tsv``). P3 to P5 (the
DrawingML and SpreadsheetML probes) are Phase C's and not copied.

The oracle is docx4j's ``ThemeColourProbesTest``, ``ThemeColorsTest``,
``ThemeColourOutputTest`` and ``WordColorsNamesTest`` at ``a9e0303ff``, as
docx4j-core-ts's ``theme-colors.test.mjs`` ported them at ``76fb45d``: each
probe case found by its label, resolved through a ``PropertyResolver`` over the
probe, and compared with Word's reading at docx4j's own tallies, family by
family (exact, within one unit, further).
"""

from __future__ import annotations

import hashlib
import math
import re
import subprocess
from collections import Counter
from pathlib import Path
from typing import Any

import pytest

from docx4j_py import create_package, load
from docx4j_py.fragments import wml
from docx4j_py.model import values
from docx4j_py.model.color import ResolvedColor, ThemeColors, WordColors
from docx4j_py.model.color import tables as color_tables
from docx4j_py.model.color import word_colors as word_colors_module
from docx4j_py.model.content.font import Font
from docx4j_py.traversal import element_name, iter_nodes
from docx4j_py.wml import Color, CTBorder, CTShd, P, R, U

ROOT = Path(__file__).resolve().parents[2]
VECTORS = ROOT / "tests" / "fixtures" / "theme-colors"
DOCX4J = ROOT.parent / "docx4j"
#: docx4j's commit the vectors and the oracle are read at.
AT = "a9e0303ff"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WNS = f'xmlns:w="{W}"'


def tsv(name: str) -> list[list[str]]:
    lines = (VECTORS / name).read_text(encoding="utf-8").splitlines()[1:]
    return [line.split("\t") for line in lines if line]


#: Word's reading per probe and label: page, fill, stroke, text, rules (docx4j's READINGS).
READINGS: dict[str, dict[str, list[str]]] = {}
for _row in tsv("readings.tsv"):
    _f = [*_row, "", "", "", "", "", ""]
    READINGS.setdefault(_f[0], {})[_f[1]] = _f[2:7]


def expected(probe: str, label: str, column: int) -> str | None:
    reading = READINGS[probe].get(label)
    return None if reading is None else reading[column]


class Tally:
    """Exact, within one unit, further: counted per family, failures listed (docx4j's ``Tally``)."""

    def __init__(self) -> None:
        self.families: dict[str, list[int]] = {}
        self.near: list[str] = []
        self.off: list[str] = []

    def add(self, family: str, label: str, got: ResolvedColor | None, want: str) -> None:
        counts = self.families.setdefault(family, [0, 0, 0])
        hex_ = want.split("/", 1)[0]
        if got is not None and got.hex() == hex_:
            counts[0] += 1
        elif got is not None and got.within(ResolvedColor.of_hex(hex_), 1):
            counts[1] += 1
            self.near.append(f"{label} {family}: ours {got}, Word {want}")
        else:
            counts[2] += 1
            self.off.append(f"{label} {family}: ours {got}, Word {want}")

    def count(self, column: int) -> int:
        return sum(counts[column] for counts in self.families.values())


#: docx4j's own tallies, family by family (exact, within one unit, further): ``ThemeColourProbesTest`` at
#: ``1fffc20a3`` as docx4j-core-ts pinned them (``org.docx4j.model.color``'s Word half is unchanged to
#: ``a9e0303ff`` but for ``normalise``, which no probe value exercises). P2's ``run text`` counts 21: the six
#: runs, the paragraph anchoring the shapes and the seven shapes' label paragraphs in each branch of their
#: ``mc:AlternateContent`` (core-ts CR-008 section 7.4 item 3). P2's shapes (``P2``, 7) are Phase C's.
TALLIES = {
    "P1": {
        "paragraph shading": [2, 1, 0],
        "pattern blend": [1, 0, 0],
        "run shading": [1, 0, 0],
        "run text": [8, 0, 0],
        "top border": [3, 0, 0],
        "underline": [1, 0, 0],
    },
    "P2": {"run text": [21, 0, 0]},
}


# ---------------------------------------------------------------------------
# the vectors
# ---------------------------------------------------------------------------


def test_the_vectors_are_docx4j_s_as_copied() -> None:
    """Every file ``SHA1SUMS`` lists, byte for byte, nothing else beside it, and the rows counted."""
    listed = {}
    for line in (VECTORS / "SHA1SUMS").read_text(encoding="utf-8").splitlines():
        digest, name = line.split("  ", 1)
        listed[name] = digest
    for name, digest in listed.items():
        assert hashlib.sha1((VECTORS / name).read_bytes()).hexdigest() == digest, name
    assert sorted(p.name for p in VECTORS.iterdir() if p.name != "SHA1SUMS") == sorted(listed)
    # core-ts's manifest.json at 76fb45d, for the six this phase copies
    assert listed == {
        "README.md": "1160075f9f6e50be38b7d8333cce514049399dde",
        "cases.tsv": "ac522b1c34f7ad09bd17bf72771524ba6a8d11a0",
        "readings.tsv": "d69d5fe58d5669a4de59e559a460896fe25f410e",
        "tc-p1-precedence.docx": "f6d75415b8f439de0c4d28a83b5fb6bc53f8b6de",
        "tc-p2-colour-map.docx": "a35e1797850bc81fcbde4f7018bdf0d41e61e4ef",
        "word-tint-shade.tsv": "834365f4878d19638e0124aaeda097de5523f2e2",
    }
    assert len(tsv("cases.tsv")) == len(tsv("readings.tsv")) == 1068
    assert len(tsv("word-tint-shade.tsv")) == 117


def _docx4j_blob(path: str) -> str:
    if not (DOCX4J / ".git").exists():
        pytest.skip("../docx4j is not checked out")
    try:
        return subprocess.run(
            ["git", "-C", str(DOCX4J), "rev-parse", f"{AT}:{path}"], check=True, capture_output=True, text=True
        ).stdout.strip()
    except subprocess.CalledProcessError:
        pytest.skip(f"../docx4j has no {path} at {AT}")


def test_each_vector_is_docx4j_s_blob_at_the_commit() -> None:
    """Each copied file hashes to its git blob in docx4j's ``theme-colors/`` at ``a9e0303ff``."""
    for name in sorted(p.name for p in VECTORS.iterdir() if p.name != "SHA1SUMS"):
        data = (VECTORS / name).read_bytes()
        blob = hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest()
        assert blob == _docx4j_blob(f"docx4j-core-tests/src/test/resources/theme-colors/{name}"), name


def test_word_tint_and_shade_against_the_corpora() -> None:
    """docx4j ``wordTintAndShadeAgainstTheCorpora``: within one unit of all 117, 76 exact (docx4j's and core-ts's)."""
    exact = near = 0
    far = []
    for base, op, byte, word in tsv("word-tint-shade.tsv"):
        got = WordColors.luminance(ResolvedColor.of_hex(base), int(byte, 16) / 255.0, op == "tint")
        if got.hex() == word:
            exact += 1
        elif got.within(ResolvedColor.of_hex(word), 1):
            near += 1
        else:
            far.append(f"{base} {op} {byte} {word} -> {got}")
    assert far == []
    assert [exact, near] == [76, 41]
    # with the written value beside it (a current cache), Word's own value comes back exactly
    theme = ThemeColors.of_theme(theme_with(accent1="5B9BD5"))
    assert WordColors.resolve("2E74B5", "accent1", None, "BF", theme).hex() == "2E74B5"


def theme_with(**slots: str) -> Any:
    """An Office 2013 ``a:theme`` with some slots replaced, parsed."""
    from docx4j_py.dml.main import Theme
    from docx4j_py.runtime import parser

    values_ = {
        "dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6", "accent1": "4472C4", "accent2": "ED7D31",
        "accent3": "A5A5A5", "accent4": "FFC000", "accent5": "4472C4", "accent6": "70AD47", "hlink": "0563C1",
        "folHlink": "954F72", **slots,
    }  # fmt: skip
    body = "".join(f'<a:{k}><a:srgbClr val="{v}"/></a:{k}>' for k, v in values_.items())
    xml = (
        f'<a:theme xmlns:a="{A}" name="t"><a:themeElements><a:clrScheme name="t">{body}</a:clrScheme>'
        '<a:fontScheme name="t"><a:majorFont><a:latin typeface="A"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        '<a:minorFont><a:latin typeface="A"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
        "<a:fmtScheme name=\"t\"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>"
        "</a:themeElements></a:theme>"
    )
    return parser().from_string(xml, Theme)


# ---------------------------------------------------------------------------
# the probes, each case through a PropertyResolver
# ---------------------------------------------------------------------------


def _text(p: Any) -> str:
    """Every ``w:t`` under the paragraph, text boxes and both branches of a kept ``mc:AlternateContent`` too.

    docx4j's ``text(p)``: ``getElementsByTagNameNS(W, "t")`` over the DOM.
    """
    return "".join(
        node.value or ""
        for node in iter_nodes(p, mce="all")
        if element_name(node) == f"{{{W}}}t" and hasattr(node, "value")
    )


def _paragraphs(package: Any) -> list[P]:
    """Every ``w:p`` of the main document part in document order, both branches of each ``mc:AlternateContent``."""
    document = package.main_document_part.contents
    return [node for node in iter_nodes(document, mce="all") if isinstance(node, P)]


def _first_run(p: P) -> R | None:
    """The paragraph's first ``w:r`` child (docx4j's ``child(p, W, "r")``)."""
    return next((item for item in p.content if isinstance(item, R)), None)


def test_p1_precedence() -> None:
    """docx4j ``p1Precedence``: a theme reference wins over the value beside it; tint and shade in HSL luminance.

    Each case through the probe's ``PropertyResolver``: the run's, the
    paragraph's effective properties and ``resolve_color``, ``resolve_fill``
    and ``resolve_pattern_color``. A family is counted where the paragraph
    states it, as docx4j's test counts. Exact but P1-10, a paragraph
    shading's ``themeFillTint`` whose written value is unusable (``FF0000``):
    Word drew ``F6C5AC``, one unit below the computation that gives its three
    run cases exactly (docx4j CR-034 section 11).
    """
    probe = "tc-p1-precedence.docx"
    package = load(VECTORS / probe)
    resolver = package.get_property_resolver()
    tally = Tally()
    labels = []
    for p in _paragraphs(package):
        text = _text(p)
        if not text.startswith("P1-"):
            continue
        label = text[:5]
        labels.append(label)
        run = _first_run(p)
        r_pr = None if run is None else run.r_pr
        p_pr = p.p_pr
        effective_r = resolver.get_effective_rpr(r_pr, p_pr)
        effective_p = resolver.get_effective_ppr(p_pr)
        own_shd = None if p_pr is None else p_pr.shd
        if r_pr is not None and r_pr.color is not None:
            tally.add("run text", label, resolver.resolve_color(effective_r.color), expected(probe, label, 3))
        if own_shd is not None and own_shd.val.value == "clear":
            tally.add("paragraph shading", label, resolver.resolve_fill(effective_p.shd), expected(probe, label, 1))
        if own_shd is not None and own_shd.val.value == "pct25":
            # the pattern's colour, then Word's 25% blend of it over the fill (Java's Math.round)
            pattern = resolver.resolve_pattern_color(effective_p.shd)
            fill = resolver.resolve_fill(effective_p.shd)
            blend = ResolvedColor.rgb(
                *(math.floor(0.25 * a + 0.75 * b + 0.5) for a, b in zip(
                    (pattern.red, pattern.green, pattern.blue), (fill.red, fill.green, fill.blue), strict=True
                ))
            )
            tally.add("pattern blend", label, blend, expected(probe, label, 1))
        if r_pr is not None and r_pr.shd is not None:
            tally.add("run shading", label, resolver.resolve_fill(effective_r.shd), expected(probe, label, 1))
        if p_pr is not None and p_pr.p_bdr is not None and p_pr.p_bdr.top is not None:
            top = resolver.resolve_color(effective_p.p_bdr.top)
            tally.add("top border", label, top, expected(probe, label, 4).split(" ")[0])
        if r_pr is not None and r_pr.u is not None:
            tally.add("underline", label, resolver.resolve_color(effective_r.u), expected(probe, label, 4).split(" ")[0])
    assert labels == [f"P1-{n:02d}" for n in range(1, 18)]
    assert tally.families == TALLIES["P1"]
    assert tally.off == []
    assert sum(tally.count(c) for c in range(3)) == 17
    assert tally.near == ["P1-10 paragraph shading: ours F6C6AD, Word F6C5AC"]


def test_p2_colour_map() -> None:
    """docx4j ``p2ColourMap``'s WordprocessingML half: ``w:clrSchemeMapping`` (light and dark swapped) applied.

    Every ``w:p`` whose text starts ``P2-``, as docx4j counts them (21), its
    first run's effective colour through the resolver; automatic text is
    black. Every case exact. The shapes (``P2-07`` to ``P2-13``'s fills,
    DrawingML ``a:schemeClr``) are Phase C's.
    """
    probe = "tc-p2-colour-map.docx"
    package = load(VECTORS / probe)
    resolver = package.get_property_resolver()
    tally = Tally()
    for p in _paragraphs(package):
        text = _text(p)
        if not text.startswith("P2-"):
            continue
        label = text[:5]
        run = _first_run(p)
        effective = resolver.get_effective_rpr(None if run is None else run.r_pr, p.p_pr)
        colour = resolver.resolve_color(effective.color)
        tally.add("run text", label, ResolvedColor.BLACK if colour is None else colour, expected(probe, label, 3))
    assert tally.families == TALLIES["P2"]
    assert tally.off == []
    # the map itself: t1 and bg1, t2 and bg2 swapped
    mapping = resolver.theme_colors.mapping
    assert [mapping[k] for k in ("bg1", "tx1", "bg2", "tx2")] == ["dk1", "lt1", "dk2", "lt2"]


def test_font_color_is_what_word_drew_for_every_p1_and_p2_paragraph_s_text() -> None:
    """core-ts's: ``Font.color`` through the content API, ``w:clrSchemeMapping`` included (17 and 6 paragraphs).

    The shaded and bordered paragraphs' text is automatic, which Word draws
    black and this reads ``""`` (CR-018 rule 3; the open question of CR-009
    section 12.9), so only a text colour Word drew other than by automatic is
    compared as ``#RRGGBB``.
    """
    for probe, count in (("tc-p1-precedence.docx", 17), ("tc-p2-colour-map.docx", 6)):
        package = load(VECTORS / probe)
        n = 0
        for paragraph in package.body.paragraphs:
            label = paragraph.text[:5]
            want = expected(probe, label, 3) if re.fullmatch(r"P[12]-\d\d", label) else None
            if not want:
                continue
            got = paragraph.font.color
            stated = paragraph.runs[0].r_pr is not None and paragraph.runs[0].r_pr.color is not None
            assert got == (f"#{want}" if stated else ""), f"{probe} {label}"
            n += 1
        assert n == count, probe


# ---------------------------------------------------------------------------
# which theme and map a docx takes (docx4j ThemeColorsTest, its Word half)
# ---------------------------------------------------------------------------


def test_new_docx_has_the_office_2023_theme() -> None:
    """docx4j ``newDocxHasTheOffice2023Theme``."""
    t = ThemeColors.of(create_package())
    assert t.scheme_color("accent1").hex() == "156082"
    assert t.word_theme_color("text1").hex() == "000000"
    assert t.scheme_color("bg1").hex() == "FFFFFF"
    assert t.word_theme_color("none") is None


def test_no_theme_part_takes_the_default_theme() -> None:
    """docx4j ``noThemePartTakesTheDefaultTheme``: the bundled theme ``pkg.default_theme`` names.

    The 2013 and 2007 settings both have Office 2013's colours (docx4j CR-034 D2).
    """
    package = create_package()
    main = package.main_document_part
    main.relationships_part.remove_part(package.theme_part.part_name)
    main.theme_part = None
    assert package.theme_part is None
    assert ThemeColors.of(package).scheme_color("accent1").hex() == "156082"
    for setting in ("2013", "2007"):
        package.default_theme = setting
        assert ThemeColors.of(package).scheme_color("accent1").hex() == "4472C4", setting
        resolver = package.get_property_resolver(refresh=True)
        assert resolver.theme_colors.slot("accent1").hex() == "4472C4"
        # core-ts's: Office 2013's accent1 shaded BF, the cache an Office 2010 value
        color = Color(val=b"\x36\x5f\x91", theme_color="accent1", theme_shade="BF")
        assert resolver.resolve_color(color).hex() == "2F5496"


def test_the_theme_part_and_the_settings_are_read_without_being_unmarshalled() -> None:
    """``ThemeColors.of`` reads both parts with ``read_contents()``: a save is byte for byte."""
    package = load(VECTORS / "tc-p2-colour-map.docx")
    theme, settings = package.theme_part, package.document_settings_part
    before = (theme.xml, settings.xml)
    assert package.get_property_resolver().theme_colors.word_theme_color("text1").hex() == "FFFFFF"
    assert theme.is_unmarshalled is False and settings.is_unmarshalled is False
    assert (theme.xml, settings.xml) == before


# ---------------------------------------------------------------------------
# through the PropertyResolver and Font (docx4j ThemeColourOutputTest)
# ---------------------------------------------------------------------------


def output_doc() -> Any:
    """docx4j's fixture: a new document in the Office 2023 theme whose stale values are Office 2007's accent1, 4F81BD.

    Office 2023's accent1 is 156082, accent2 E97132, accent4 0F9ED5, accent5
    A02B93, accent6 4EA72E. Each paragraph's text names its case.
    """
    package = create_package()
    styles = package.style_definitions_part.contents
    for style_id, pr in (
        ("StaleColour", '<w:rPr><w:color w:val="4F81BD" w:themeColor="accent1"/></w:rPr>'),
        ("ThemeOnly", '<w:rPr><w:color w:themeColor="accent2"/></w:rPr>'),
        ("CurrentCache", '<w:rPr><w:color w:val="156082" w:themeColor="accent1"/></w:rPr>'),
    ):
        styles.style.append(
            wml(
                f'<w:style {WNS} w:type="paragraph" w:styleId="{style_id}"><w:name w:val="{style_id}"/>'
                f'<w:basedOn w:val="Normal"/>{pr}</w:style>'
            )
        )
    package.get_property_resolver().refresh()
    body = package.main_document_part.contents.body
    for p in (
        '<w:pPr><w:pStyle w:val="StaleColour"/></w:pPr><w:r><w:t>STALE</w:t></w:r>',
        '<w:pPr><w:pStyle w:val="StaleColour"/></w:pPr><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>RED</w:t></w:r>',
        '<w:pPr><w:pStyle w:val="ThemeOnly"/></w:pPr><w:r><w:t>THEMEONLY</w:t></w:r>',
        '<w:pPr><w:pStyle w:val="CurrentCache"/></w:pPr><w:r><w:t>CURRENT</w:t></w:r>',
        '<w:r><w:rPr><w:color w:val="auto" w:themeColor="accent5"/></w:rPr><w:t>AUTO</w:t></w:r>',
        (
            '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="auto" w:themeFill="accent2"/></w:pPr>'
            "<w:r><w:t>SHADEDAUTO</w:t></w:r>"
        ),
        "<w:r><w:t>UNBORDERED</w:t></w:r>",
        (
            '<w:pPr><w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="4F81BD" w:themeColor="accent6"/>'
            "</w:pBdr></w:pPr><w:r><w:t>BORDERED</w:t></w:r>"
        ),
        '<w:r><w:rPr><w:u w:val="single" w:color="4F81BD" w:themeColor="accent4"/></w:rPr><w:t>UNDERLINED</w:t></w:r>',
    ):
        body.content.append(wml(f"<w:p {WNS}>{p}</w:p>"))
    return package


def _effective_r_pr(package: Any, i: int) -> Any:
    p = package.body.paragraphs[i].element
    return package.get_property_resolver().get_effective_rpr(_first_run(p).r_pr, p.p_pr)


def test_resolver_gives_the_colour_word_draws() -> None:
    """docx4j ``resolverGivesTheColourWordDraws``, with core-ts's underline and automatic cases."""
    package = output_doc()
    resolver = package.get_property_resolver()
    hexes = [resolver.resolve_color(_effective_r_pr(package, i).color).hex() for i in range(5)]
    assert hexes == ["156082", "FF0000", "E97132", "156082", "A02B93"]  # STALE RED THEMEONLY CURRENT AUTO
    shaded = resolver.get_effective_ppr(package.body.paragraphs[5].element.p_pr)
    assert resolver.resolve_fill(shaded.shd).hex() == "E97132"
    assert resolver.resolve_pattern_color(shaded.shd) is None, "automatic"
    bordered = resolver.get_effective_ppr(package.body.paragraphs[7].element.p_pr)
    assert resolver.resolve_color(bordered.p_bdr.top).hex() == "4EA72E"
    assert resolver.resolve_color(_effective_r_pr(package, 8).u).hex() == "0F9ED5"
    assert resolver.resolve_color(_effective_r_pr(package, 6).color) is None, "no colour anywhere: automatic"
    assert resolver.theme_colors is resolver.theme_colors is resolver.get_theme_colors(), "made once"
    # refresh() reads the theme again
    before = resolver.theme_colors
    package.get_property_resolver(refresh=True)
    assert resolver.theme_colors is not before


def test_a_direct_value_replaces_the_style_s_theme_reference() -> None:
    """docx4j ``aDirectValueReplacesTheStylesThemeReference``."""
    c = _effective_r_pr(output_doc(), 1).color
    assert values.colour(c.val) == "FF0000"
    assert c.theme_color is None


def test_a_theme_colour_with_no_value_is_inherited() -> None:
    """docx4j ``aThemeColourWithNoValueIsInherited``."""
    c = _effective_r_pr(output_doc(), 2).color
    assert c.theme_color.value == "accent2"


def test_font_color_reads_the_colour_word_draws_and_the_direct_view_w_val() -> None:
    """core-ts's: the effective view the drawn colour, the direct view ``w:val`` as written (its D2).

    Unlike core-ts, an automatic colour reads ``""`` (none stated) or ``"auto"``
    here, not ``#000000`` (CR-018 rule 3; CR-009 section 12.9's question).
    """
    package = output_doc()
    ps = package.body.paragraphs
    assert [ps[i].font.color for i in (0, 1, 2, 3, 4, 6)] == ["#156082", "#FF0000", "#E97132", "#156082", "#A02B93", ""]
    whole = [ps[i].get_range("Whole").font.color for i in (0, 1, 2, 3, 4, 6)]
    assert whole == ["#156082", "#FF0000", "#E97132", "#156082", "#A02B93", ""]
    assert [ps[i].get_font(direct=True).color for i in (0, 1, 4)] == ["", "#FF0000", "auto"]
    # across paragraphs whose runs' colours differ, '' (MixedFont); where they agree, the colour drawn
    assert ps[0].get_range("Whole").expand_to(ps[1].get_range("Whole")).font.color == ""
    assert ps[0].get_range("Whole").expand_to(ps[3].get_range("Whole")).font.color == ""
    assert ps[3].get_range("Whole").expand_to(ps[3].get_range("Whole")).font.color == "#156082"


def test_to_dict_gives_the_theme_attributes_as_written() -> None:
    """``Font.to_dict()``'s ``theme_color``, ``theme_tint`` and ``theme_shade``: effective and direct, as written."""
    package = output_doc()
    ps = package.body.paragraphs
    stale = ps[0].font.to_dict()
    assert (stale["color"], stale["theme_color"], stale["theme_tint"], stale["theme_shade"]) == (
        "#156082", "accent1", None, None,
    )  # fmt: skip
    assert {k: ps[0].get_font(direct=True).to_dict()[k] for k in ("theme_color", "theme_tint", "theme_shade")} == {
        "theme_color": None, "theme_tint": None, "theme_shade": None,
    }  # fmt: skip
    auto = ps[4].get_font(direct=True).to_dict()
    assert (auto["color"], auto["theme_color"]) == ("auto", "accent5")
    heading = create_package()
    first = heading.body.insert_paragraph("Title", style="Heading 1")
    second = heading.body.insert_paragraph("Another", style="Heading 1")
    read = first.font.to_dict()
    assert (read["color"], read["theme_color"], read["theme_shade"]) == ("#0F4761", "accent1", "BF")
    # a range across paragraphs: where the runs agree, the attributes; else None
    across = first.get_range("Whole").expand_to(second.get_range("Whole")).font.to_dict()
    assert (across["color"], across["theme_color"], across["theme_shade"]) == ("#0F4761", "accent1", "BF")
    across = ps[0].get_range("Whole").expand_to(ps[2].get_range("Whole")).font.to_dict()
    assert (across["color"], across["theme_color"]) == ("", None)


def test_a_colour_written_through_font_has_no_theme_reference() -> None:
    """Setting ``Font.color`` writes ``w:val`` alone, so the run's own value replaces a style's reference."""
    package = create_package()
    paragraph = package.body.insert_paragraph("Title", style="Heading 1")
    paragraph.font.color = "#FF0000"
    assert paragraph.font.color == "#FF0000"
    assert paragraph.font.to_dict()["theme_color"] is None


# ---------------------------------------------------------------------------
# WordColors' readings of a value (docx4j WordColorsNamesTest)
# ---------------------------------------------------------------------------


def test_names_and_hex_spellings_in_every_colour() -> None:
    """docx4j ``namesAndHexSpellingsInEveryColour``."""
    assert WordColors.of(Color(val="red"), None).hex() == "FF0000"
    assert WordColors.of(U(color="Navy"), None).hex() == "000080"
    shd = CTShd(fill="Aqua", color=" #ff0000 ")
    assert WordColors.fill(shd, None).hex() == "00FFFF"
    assert WordColors.pattern(shd, None).hex() == "FF0000"
    assert WordColors.of(CTBorder(color="navy"), None).hex() == "000080"


def test_a_rejected_value_is_automatic() -> None:
    """docx4j ``aRejectedValueIsAutomatic``."""
    shd = CTShd(fill="F00")
    assert WordColors.fill(shd, None) is None
    assert WordColors.of(CTBorder(color="bogus"), None) is None
    shd.fill = "auto"
    assert WordColors.fill(shd, None) is None


def test_normalise_and_stated_are_the_one_reader_s() -> None:
    """``WordColors.normalise`` and ``stated`` are :mod:`docx4j_py.model.values`' (CR-018), not a second table."""
    spellings = ["FF0000", "ff0000", " #FF0000 ", "red", "DarkRed", "auto", "Auto", "F00", "grey", "", None]
    for value in spellings:
        read = values.colour(value)
        assert WordColors.normalise(value) == (None if read in (None, values.AUTO) else read), value
        assert WordColors.stated(value) is values.states_colour(value), value
    assert word_colors_module.colour is values.colour
    assert word_colors_module.states_colour is values.states_colour
    names = {
        name
        for name in dir(word_colors_module)
        if not name.startswith("__") and isinstance(getattr(word_colors_module, name), dict)
    }
    assert not names, "no colour table of its own"


def _docx4j_source(path: str) -> str:
    if not (DOCX4J / ".git").exists():
        pytest.skip("../docx4j is not checked out")
    try:
        return subprocess.run(
            ["git", "-C", str(DOCX4J), "show", f"{AT}:{path}"], check=True, capture_output=True, text=True
        ).stdout
    except subprocess.CalledProcessError:
        pytest.skip(f"../docx4j has no {path} at {AT}")


def _java_table(source: str, name: str) -> dict[str, str] | None:
    body = re.search(rf"Map<String, String> {name} = table\((.*?)\);", source, re.DOTALL)
    if body is None:
        return None
    items = re.findall(r'"([^"]+)"', body.group(1))
    return dict(zip(items[::2], items[1::2], strict=True))


def test_the_tables_are_docx4j_s() -> None:
    """The colour names Word reads are docx4j's ``WordColorValues.HTML`` at ``a9e0303ff`` (no longer ``ColorTables``');
    ``SYSTEM`` is ``ColorTables.SYSTEM``."""
    color_tables_java = _docx4j_source("docx4j-core/src/main/java/org/docx4j/model/color/ColorTables.java")
    word_color_values = _docx4j_source("docx4j-generated-objects/src/main/java/org/docx4j/wml/WordColorValues.java")
    assert _java_table(color_tables_java, "HTML") is None
    assert _java_table(word_color_values, "HTML") == values.NETSCAPE_COLOURS
    assert len(values.NETSCAPE_COLOURS) == 140
    assert _java_table(color_tables_java, "SYSTEM") == dict(color_tables.SYSTEM)


# ---------------------------------------------------------------------------
# the units
# ---------------------------------------------------------------------------


def test_resolved_color() -> None:
    """``ResolvedColor``: OOXML's alpha, ``of_hex``, ``hex``, ``within``, the clamps."""
    c = ResolvedColor.of_hex("#0f4761")
    assert (c.red, c.green, c.blue, c.alpha, c.is_opaque) == (15, 71, 97, 100000, True)
    assert c.hex() == "0F4761" and str(c) == "0F4761"
    assert str(c.with_alpha(50000)) == "0F4761/50000"
    assert ResolvedColor.rgb(-3, 300, 7) == ResolvedColor(0, 255, 7)
    for bad in ("auto", None, "F00", "GG0000", "+FFFFF", " 0F4761", b"\x0f\x47\x61"):
        assert ResolvedColor.of_hex(bad) is None, bad
    assert ResolvedColor.BLACK.hex() == "000000" and ResolvedColor.WHITE.hex() == "FFFFFF"
    assert c.within(ResolvedColor.rgb(16, 70, 98), 1) and not c.within(ResolvedColor.rgb(17, 71, 97), 1)


def test_none_gives_none_and_no_reference_gives_the_written_value() -> None:
    theme = ThemeColors.of_theme(None, "2023")
    assert WordColors.resolve("FF0000", "none", None, None, theme).hex() == "FF0000"
    assert WordColors.resolve("auto", "none", None, None, theme) is None
    assert WordColors.resolve(None, None, None, None, theme) is None
    assert WordColors.resolve("FF0000", "accent1", None, None, None).hex() == "FF0000", "no theme"
    # a reference outside ST_ThemeColor names no colour: the written value stands
    assert WordColors.resolve("FF0000", "Accent1", None, None, theme).hex() == "FF0000"
    assert WordColors.of(None, theme) is None and WordColors.fill(None, theme) is None
    with pytest.raises(TypeError):
        WordColors.of(CTShd(), theme)


def test_the_cache_rule() -> None:
    """Word's written value where it is within one unit of the computation; two units off, the computation."""
    theme = ThemeColors.of_theme(None, "2023")  # accent1 156082; shaded BF, 0F4761
    assert WordColors.resolve("0F4761", "accent1", None, "BF", theme).hex() == "0F4761"
    assert WordColors.resolve("104862", "accent1", None, "BF", theme).hex() == "104862", "one unit off: Word's"
    assert WordColors.resolve("114963", "accent1", None, "BF", theme).hex() == "0F4761", "two units off: replaced"
    assert WordColors.resolve("0F4762", "accent1", None, None, theme).hex() == "156082"


def test_the_tint_wins_over_the_shade() -> None:
    theme = ThemeColors.of_theme(None, "2023")
    tint = WordColors.resolve(None, "accent1", "99", None, theme)
    assert tint.hex() == "45B0E1"  # P1-02's
    assert WordColors.resolve(None, "accent1", "99", "BF", theme) == tint
    assert WordColors.resolve(None, "accent1", "zz", "BF", theme).hex() == "0F4761", "an unusable tint: the shade"
    assert WordColors.resolve(None, "accent1", "0x99", None, theme).hex() == "156082", "Java's parseInt, not int()"


def test_the_word_colour_scheme_mapping_swaps_the_slots() -> None:
    """``with_word_mapping``: ``w:clrSchemeMapping``'s names (``t1``, ``dark1``) to the scheme's slots."""
    theme = ThemeColors.of_theme(None, "2023")
    mapping = wml(
        f'<w:clrSchemeMapping {WNS} w:bg1="dark1" w:t1="light1" w:bg2="dark2" w:t2="light2" w:accent1="accent2" '
        'w:accent2="accent1" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" '
        'w:hyperlink="followedHyperlink" w:followedHyperlink="hyperlink"/>'
    )
    swapped = theme.with_word_mapping(mapping)
    assert [swapped.word_theme_color(n).hex() for n in ("text1", "background1", "text2", "background2")] == [
        "FFFFFF", "000000", "E8E8E8", "0E2841",
    ]  # fmt: skip
    # w:themeColor's accents and hyperlinks name their slots directly, as docx4j's wordThemeColor does
    assert swapped.word_theme_color("accent1").hex() == "156082"
    assert swapped.word_theme_color("hyperlink").hex() == "467886"
    assert swapped.scheme_color("accent1").hex() == "E97132", "a DrawingML schemeClr goes through the map"
    assert theme.word_theme_color("text1").hex() == "000000", "immutable: the original keeps its map"
    assert theme.with_word_mapping(None) is theme


@pytest.mark.filterwarnings("ignore::docx4j_xsdata.exceptions.ConverterWarning")  # sysClr "bogus", kept on purpose
def test_a_slot_phase_c_resolves_is_left_to_the_written_value() -> None:
    """The seam: a scheme slot other than a plain ``a:srgbClr`` or ``a:sysClr`` is not resolved before Phase C."""
    from docx4j_py.dml.main import Theme
    from docx4j_py.runtime import parser

    xml = (
        f'<a:theme xmlns:a="{A}" name="t"><a:themeElements><a:clrScheme name="t">'
        '<a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1>'
        '<a:lt1><a:sysClr val="bogus" lastClr="123456"/></a:lt1>'
        '<a:accent1><a:srgbClr val="4472C4"><a:lumMod val="75000"/></a:srgbClr></a:accent1>'
        '<a:accent2><a:prstClr val="red"/></a:accent2>'
        "</a:clrScheme></a:themeElements></a:theme>"
    )
    theme = ThemeColors.of_theme(parser().from_string(xml, Theme))
    assert theme.slot("dk1").hex() == "000000", "the machine's colour for the name, not lastClr"
    assert theme.slot("lt1").hex() == "123456", "an unknown name: lastClr"
    assert theme.slot("accent1") is None and theme.slot("accent2") is None
    assert WordColors.resolve("2F5496", "accent1", None, None, theme).hex() == "2F5496"


def test_every_colour_family_of_a_p1_probe_counts() -> None:
    """The tallies' families are docx4j's six, no more (a guard on :data:`TALLIES`)."""
    notes = Counter(row[2] for row in tsv("cases.tsv") if row[0] == "tc-p1-precedence.docx")
    assert notes == {
        "run text": 8, "paragraph shading": 4, "paragraph top border": 3, "underline": 1, "run shading": 1,
    }  # fmt: skip


def test_a_font_over_a_detached_run_reads_its_own_colour() -> None:
    """A font with no package behind it reads direct: ``w:val`` and the attributes as written."""
    run = wml(f'<w:r {WNS}><w:rPr><w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/></w:rPr></w:r>')
    font = Font(lambda: [run])
    assert font.color == "#365F91"
    assert (font.to_dict()["theme_color"], font.to_dict()["theme_shade"]) == ("accent1", "BF")
