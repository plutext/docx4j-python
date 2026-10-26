"""Integer attributes read as Word 365 reads them: CR-021, over docx4j CR-065's probe and Word's saves.

The fixtures are ``tests/fixtures/integers/`` (``README.md``, ``SHA1SUMS``): the
probe's families and Word 365's save of each family it opened, rounds two to four,
2026-10-26. Every case is a labelled paragraph (``"B20 w:ind/@w:left=..."``) whose
own properties, or the table or section that follows it, or (family ``I``) the
styles, settings or header, carry the value; :data:`LOCATIONS` says where.

**The oracle.** Word's save of a family holds Word's reading of each value, an
integer, or nothing where it equals the style's (``README.md``: ``A08``, ``A17``,
``B25``, ``F41``, :data:`OMITTED`). A family Word refused has no save, and each of
its values reads None. So for every case :func:`values.integer` of the source's
value is Word's integer, or None where Word refused the document
(:func:`test_every_case_reads_as_word_reads_it`), and the readers over the probe
read as over Word's save.
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

from docx4j_py.model.values import integer

ROOT = Path(__file__).resolve().parents[2]
INTEGERS = ROOT / "tests" / "fixtures" / "integers"

# the probe holds values the int converter refuses ("720.5", ""): each is kept and warned about
pytestmark = pytest.mark.filterwarnings("ignore::docx4j_xsdata.exceptions.ConverterWarning")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def _w(local: str) -> str:
    return f"{{{W}}}{local}"


# ---------------------------------------------------------------------------
# the fixtures and the cases
# ---------------------------------------------------------------------------


def test_the_fixtures_are_as_copied() -> None:
    """Every file ``SHA1SUMS`` lists, byte for byte, and nothing else beside it."""
    listed = {}
    for line in (INTEGERS / "SHA1SUMS").read_text(encoding="utf-8").splitlines():
        digest, name = line.split("  ", 1)
        listed[name] = digest
    for name, digest in listed.items():
        assert hashlib.sha1((INTEGERS / name).read_bytes()).hexdigest() == digest, name
    assert sorted(p.name for p in INTEGERS.iterdir() if p.name != "SHA1SUMS") == sorted(listed)
    assert len([n for n in listed if n.endswith("-word365.docx")]) == 14
    assert len([n for n in listed if n.endswith(".docx") and not n.endswith("-word365.docx")]) == 22


def _cases() -> dict[str, str]:
    """``cases.txt``: case id to its description (``w:spacing/@w:before="120.5"``)."""
    out = {}
    for line in (INTEGERS / "cases.txt").read_text(encoding="utf-8").splitlines():
        _family, rest = line.split("\t", 1)
        case, description = rest.split(" ", 1)
        out[case] = description
    return out


CASES = _cases()

#: The smallest family file holding each case: rounds three and four split ``A``.
FAMILY = {
    **{f"A0{n}": "cr065-A1-decimals" for n in range(1, 7)},
    "A07": "cr065-A2a-trailing-dot",
    "A08": "cr065-A2b-leading-dot",
    "A09": "cr065-A2c-plus",
    "A10": "cr065-A2d-plus-decimal",
    "A11": "cr065-A2e-spaced-decimal",
    "A12": "cr065-A3-comma",
    "A13": "cr065-A4-hex",
    "A14": "cr065-A5-exponents",
    "A15": "cr065-A5-exponents",
    "A16": "cr065-A5-exponents",
    "A17": "cr065-A6-empty",
    "A18": "cr065-A7-abc",
    "A19": "cr065-A2f-spaced-integer",
    **{case: f"cr065-{case[0]}" for case in CASES if case[0] in "BCDEFGHI"},
    "Z65": "cr065-A1-decimals",  # the control, in every family
}

#: Where each case's value is: (anchor, path from it, attribute). The anchor is the label
#: paragraph (``p``), the element after it (``next``: a table, or the paragraph whose
#: ``w:sectPr`` ends the section), or a part (``styles``, ``settings``, ``header``).
LOCATIONS: dict[str, list[tuple[str, str, str]]] = {
    **{f"A{n:02}": [("p", "w:pPr/w:spacing", "before")] for n in range(1, 20)},
    **{f"B{n}": [("p", "w:pPr/w:ind", "left")] for n in range(20, 28)},
    "C28": [("p", "w:pPr/w:spacing", "after")],
    "C29": [("p", "w:pPr/w:spacing", "after")],
    "C30": [("p", "w:pPr/w:spacing", "line")],
    "C31": [("p", "w:pPr/w:spacing", "line")],
    "C32": [("p", "w:pPr/w:spacing", "line")],
    "C33": [("p", "w:pPr/w:ind", "hanging")],
    "C34": [("p", "w:pPr/w:ind", "firstLine")],
    "D35": [("p", "w:pPr/w:pBdr/w:top", "sz")],
    "D36": [("p", "w:pPr/w:pBdr/w:top", "space")],
    "D37": [("p", "w:pPr/w:pBdr/w:bottom", "space")],
    "D38": [("p", "w:pPr/w:pBdr/w:between", "space")],
    "E39": [("p", "w:pPr/w:tabs/w:tab", "pos")],
    "E40": [("p", "w:pPr/w:tabs/w:tab", "pos")],
    "F41": [("p", "w:r/w:rPr/w:sz", "val")],
    "F42": [("p", "w:r/w:rPr/w:w", "val")],
    "F43": [("p", "w:r/w:rPr/w:spacing", "val")],
    "F44": [("p", "w:r/w:rPr/w:position", "val")],
    **{f"G{n}": [("next", "w:tblPr/w:tblW", "w")] for n in (45, 46, 47, 48, 55)},
    "G49": [("next", "w:tblPr/w:tblInd", "w")],
    "G50": [("next", "w:tblGrid/w:gridCol[1]", "w"), ("next", "w:tblGrid/w:gridCol[2]", "w")],
    "G51": [("next", "w:tblGrid/w:gridCol[1]", "w")],
    "G52": [("next", "w:tr/w:trPr/w:trHeight", "val")],
    "G53": [("next", "w:tr/w:trPr/w:trHeight", "val")],
    "G54": [("next", "w:tr/w:tc[1]/w:tcPr/w:tcW", "w")],
    "H56": [("next", "w:pPr/w:sectPr/w:pgSz", "w"), ("next", "w:pPr/w:sectPr/w:pgSz", "h")],
    "H57": [("next", "w:pPr/w:sectPr/w:pgSz", "w"), ("next", "w:pPr/w:sectPr/w:pgSz", "h")],
    "H58": [("next", "w:pPr/w:sectPr/w:pgMar", "top")],
    "H59": [("next", "w:pPr/w:sectPr/w:pgMar", "top")],
    "H60": [("next", "w:pPr/w:sectPr/w:pgMar", name) for name in ("header", "footer", "gutter")],
    "H61": [("next", "w:pPr/w:sectPr/w:pgMar", "top")],
    "I62": [("styles", "w:style[@w:styleId='ProbeStyle']/w:pPr/w:ind", "left")],
    "I63": [("settings", "w:defaultTabStop", "val")],
    "I64": [("header", "w:p/w:pPr/w:spacing", "before")],
    "Z65": [("p", "w:pPr/w:spacing", "before")],
}

#: Values Word reads but does not write, because each equals the style's after truncation
#: (``README.md``; CR-021 section 3.1): what Word read, which this package's save does write (D7).
OMITTED: dict[tuple[str, str], int] = {
    ("A08", "before"): 0,  # ".5"
    ("A17", "before"): 0,  # ""
    ("B25", "left"): 0,  # "-0.5"
    ("F41", "val"): 22,  # "22.5": the document default's w:sz
}

#: ``w:gridCol``: Word recomputes the grid on open, so its save shows no reading (CR-021 section 3.1).
GRID = frozenset({"G50", "G51"})

#: The families Word refused, so with no save (rounds two to four).
REFUSED = frozenset(
    {"cr065-A2e-spaced-decimal", "cr065-A2f-spaced-integer", "cr065-A3-comma", "cr065-A4-hex",
     "cr065-A5-exponents", "cr065-A7-abc"}
)


class Family:
    """A family file, or Word's save of it, as lxml trees: the parts a case's value can be in."""

    def __init__(self, data: bytes) -> None:
        archive = zipfile.ZipFile(io.BytesIO(data))
        names = archive.namelist()
        self.document = etree.fromstring(archive.read("word/document.xml"))
        self.styles = etree.fromstring(archive.read("word/styles.xml"))
        self.settings = etree.fromstring(archive.read("word/settings.xml"))
        headers = sorted(n for n in names if n.startswith("word/header") and n.endswith(".xml"))
        self.header = etree.fromstring(archive.read(headers[0])) if headers else None

    def label(self, case: str) -> Any:
        """The body paragraph whose text starts with the case id."""
        for p in self.document.find(_w("body")).iterchildren(_w("p")):
            if "".join(p.itertext()).startswith(case):
                return p
        raise LookupError(case)

    def values(self, case: str) -> list[tuple[str, str | None]]:
        """Each of the case's attributes, and its value here (None where absent)."""
        out = []
        for anchor, path, attribute in LOCATIONS[case]:
            if anchor == "p":
                base = self.label(case)
            elif anchor == "next":
                base = self.label(case).getnext()
            else:
                base = getattr(self, anchor)
            found = base.xpath(path, namespaces=NS)
            out.append((attribute, found[0].get(_w(attribute)) if found else None))
        return out


def _family(name: str) -> Family:
    return Family((INTEGERS / f"{name}.docx").read_bytes())


def _save(name: str) -> Family | None:
    path = INTEGERS / f"{name}-word365.docx"
    return Family(path.read_bytes()) if path.exists() else None


def test_cases_txt_is_the_probe() -> None:
    """The 65 cases, each found where :data:`LOCATIONS` says, holding the value ``cases.txt`` names."""
    assert len(CASES) == 65 and set(CASES) == set(LOCATIONS) == set(FAMILY)
    for case, description in CASES.items():
        values = [value for _attribute, value in _family(FAMILY[case]).values(case)]
        assert all(value is not None for value in values), case
        for value in values:
            assert f'"{value}"' in description, (case, value, description)
    assert {name for name in set(FAMILY.values()) if not (INTEGERS / f"{name}-word365.docx").exists()} == REFUSED


# ---------------------------------------------------------------------------
# D1: values.integer
# ---------------------------------------------------------------------------

#: (value, Word's reading). The measured forms, then the unmeasured edges D1 decides.
SPELLINGS: list[tuple[Any, int | None]] = [
    ("120.5", 120), ("120.4", 120), ("121.5", 121), ("122.5", 122), ("240.0", 240),
    ("141.99999999999994", 141), ("240.", 240), (".5", 0), ("+240", 240), ("+240.5", 240),
    ("720.5", 720), ("-120.5", -120), ("-121.5", -121), ("-0.5", 0), ("-1.5", -1), ("10206.0", 10206),
    ("", 0), ("0", 0), ("240", 240), ("0240", 240), ("-240", -240),
    # refused by Word (rounds three and four)
    (" 240.5 ", None), ("240,5", None), ("0x10", None), ("1.5e2", None), ("2.5e0", None), ("1e3", None),
    ("abc", None), (" 240 ", None),
    # unmeasured, read by D1's rule (CR-021 section 4.1; Q8 adds them to phase C's Word check)
    ("2.9999999999999999", 2), (".", 0), ("+", 0), ("-", 0), ("-.", 0), ("+.5", 0),
    # never Word's form: whitespace, underscores, other scripts' digits, two dots, a trailing newline
    ("\t240\n", None), ("240\n", None), ("1_000", None), ("٢٤٠", None), ("1.2.3", None),
    ("--1", None), ("+-1", None), ("1-", None), ("1.5.", None), ("２４０", None),
    # not text
    (240, 240), (-3, -3), (0, 0), (None, None), (True, None), (False, None), (1.5, None), (b"240", None),
]


@pytest.mark.parametrize(("value", "read"), SPELLINGS)
def test_integer_reads_as_word_reads(value: Any, read: int | None) -> None:
    assert integer(value) == read
    assert integer(value) is None or type(integer(value)) is int or isinstance(value, int)


def test_integer_truncates_the_string_not_a_float() -> None:
    """``float("2.9999999999999999")`` is 3.0; the string's integer part is 2 (CR-021 section 4.1)."""
    assert float("2.9999999999999999") == 3.0
    assert integer("2.9999999999999999") == 2
    assert integer("-2.9999999999999999") == -2
    assert integer("12345678901234567890.5") == 12345678901234567890  # no 16-bit wrap (Q7)
    assert integer("9" * 5000) is None  # beyond Python's conversion limit: no exception


# ---------------------------------------------------------------------------
# the oracle: every case against Word
# ---------------------------------------------------------------------------


def test_every_case_reads_as_word_reads_it() -> None:
    """``values.integer`` of each source value is Word's integer, or None where Word refused the document."""
    checked = refused = omitted = 0
    for case in CASES:
        name = FAMILY[case]
        source = _family(name).values(case)
        saved = _save(name)
        if saved is None:
            assert name in REFUSED
            assert [integer(value) for _attribute, value in source] == [None] * len(source), case
            refused += 1
            continue
        for (attribute, value), (_same, word) in zip(source, saved.values(case), strict=True):
            if case in GRID:
                assert word is not None  # Word opened it and wrote a grid of its own
                continue
            if word is None:
                assert integer(value) == OMITTED[(case, attribute)], (case, attribute, value)
                omitted += 1
                continue
            assert integer(value) == int(word), (case, attribute, value, word)
            checked += 1
    # 8 cases in the refused families (A11 to A16, A18, A19; the three exponents among them), 4 omitted,
    # and of the 70 values (65 cases; G50, H56 and H57 two each, H60 three) 55 Word wrote
    values = sum(len(found) for found in LOCATIONS.values())
    assert (values, refused, omitted) == (70, 8, 4)
    assert checked == values - refused - omitted - 3 == 55  # less the grid's three


# ---------------------------------------------------------------------------
# D2: the readers, over the probe and over Word's save
# ---------------------------------------------------------------------------


def _load(path: Path) -> Any:
    from docx4j_py.openpackaging.api import load

    return load(path)


def _paragraph_readings(package: Any) -> dict[str, tuple[float, ...]]:
    """The content API's paragraph measurements of every labelled paragraph."""
    out = {}
    for p in package.body.paragraphs:
        label = p.text[:3]
        if label in CASES:
            out[label] = (p.left_indent, p.first_line_indent, p.space_before, p.space_after, p.line_spacing)
    return out


@pytest.mark.parametrize(
    "name",
    ["cr065-A1-decimals", "cr065-A2a-trailing-dot", "cr065-A2b-leading-dot", "cr065-A2c-plus",
     "cr065-A2d-plus-decimal", "cr065-A6-empty", "cr065-B", "cr065-C"],
)
def test_the_content_api_reads_as_word(name: str) -> None:
    """``Paragraph.left_indent``, ``first_line_indent``, ``space_before``, ``space_after``, ``line_spacing``.

    Over the probe as over Word's save: ``"720.5"`` is 36 pt, as Word shows it,
    where ``float()`` read 36.025.
    """
    probe = _paragraph_readings(_load(INTEGERS / f"{name}.docx"))
    word = _paragraph_readings(_load(INTEGERS / f"{name}-word365.docx"))
    assert probe and probe == word


def _facade_readings(path: Path) -> dict[str, Any]:
    from docx4j_py.docx import Document

    document = Document(path)
    out: dict[str, Any] = {}
    for paragraph in document.paragraphs:
        label = paragraph.text[:3]
        if label not in CASES:
            continue
        fmt = paragraph.paragraph_format
        out[label] = (
            fmt.left_indent, fmt.first_line_indent, fmt.space_before, fmt.space_after, fmt.line_spacing,
            [stop.position for stop in fmt.tab_stops],
        )
    out["sections"] = [
        (s.page_width, s.page_height, s.top_margin, s.header_distance, s.footer_distance, s.gutter)
        for s in document.sections
    ]
    out["tables"] = [
        ([cell.width for cell in table.rows[0].cells], [row.height for row in table.rows])
        for table in document.tables
    ]
    return out


@pytest.mark.parametrize("name", ["cr065-A1-decimals", "cr065-B", "cr065-C", "cr065-E", "cr065-H"])
def test_the_facade_reads_as_word(name: str) -> None:
    """The python-docx facade's lengths, over the probe as over Word's save (it raised ``ValueError`` on H).

    A value Word omits from its save (:data:`OMITTED`) the facade reads as
    stated, ``Twips(0)`` for ``B25``'s ``"-0.5"``, where Word's save states
    nothing and the facade reads None, as python-docx does.
    """
    probe = _facade_readings(INTEGERS / f"{name}.docx")
    word = _facade_readings(INTEGERS / f"{name}-word365.docx")
    for case, _attribute in OMITTED:
        if case in probe:
            assert probe.pop(case)[0] == 0 and word.pop(case)[0] is None, case
    assert probe == word


def test_the_facade_no_longer_raises() -> None:
    """CR-021 section 2.3's three ``ValueError``\\ s: Google Docs' ``"11907.0"``, SSRS's ``""``, a ``"4500.5"`` grid."""
    from docx4j_py.docx import Document
    from docx4j_py.docx.shared import Twips

    sections = Document(INTEGERS / "cr065-H.docx").sections
    assert (sections[0].page_width, sections[0].page_height) == (Twips(11907), Twips(16839))  # H56
    assert (sections[1].page_width, sections[1].page_height) == (Twips(11907), Twips(16838))  # H57
    assert sections[4].header_distance == sections[4].footer_distance == sections[4].gutter == Twips(0)  # H60
    assert sections[5].top_margin == Twips(0)  # H61, a lone top=""
    tables = Document(INTEGERS / "cr065-G.docx").tables
    assert [column.width for column in tables[5].columns] == [Twips(4500), Twips(5705)]  # G50
    assert tables[0].rows[0].cells[0].width == Twips(4500)
    assert tables[7].rows[0].height == Twips(400)  # G52


def test_the_content_api_s_tables_and_sections_read_as_word() -> None:
    from types import SimpleNamespace

    from docx4j_py.model.content.describe import describe_package
    from docx4j_py.model.content.picture import writable_width_emu
    from docx4j_py.model.content.table import writable_width
    from docx4j_py.wml import P

    for name in ("cr065-G", "cr065-G-word365"):
        tables = _load(INTEGERS / f"{name}.docx").body.tables
        assert tables[9].cell(0, 0).width == 0.0  # G54: a pct width is no fixed one
        assert tables[0].cell(0, 0).width == 4500 / 20
    grid = _load(INTEGERS / "cr065-G.docx").body.tables[5]
    assert grid.column_widths_twips() == [4500, 5705]  # G50, as Word reads each before it recomputes the grid
    # a section's page as the body's: H57's 11907.5 by 16838.5, H60's empty header, footer and gutter
    paragraphs = [p for p in _load(INTEGERS / "cr065-H.docx").main_document_part.contents.body.content if isinstance(p, P)]
    sections = [p.p_pr.sect_pr for p in paragraphs if p.p_pr is not None and p.p_pr.sect_pr is not None]
    assert sections[1].pg_sz.w == "11907.5"  # kept as read
    body = SimpleNamespace(container=SimpleNamespace(sect_pr=sections[1]))
    assert writable_width(body) == 11907 - 1440 - 1440
    assert writable_width_emu(body) == (11907 - 1440 - 1440) * 635
    page = describe_package(_load(INTEGERS / "cr065-H.docx")).page  # the final section's, read with lxml
    assert (page.width_pt, page.height_pt) == (11906 / 20, 16838 / 20)


def test_the_resolver_and_the_text_model_read_as_word() -> None:
    """A style's ``w:ind/@w:left="720.5"`` (I62) is 720 twips in the resolver's effective properties and the text model."""
    from docx4j_py.model.text import extract

    for name in ("cr065-I", "cr065-I-word365"):
        package = _load(INTEGERS / f"{name}.docx")
        resolver = package.get_property_resolver()
        effective = resolver.get_effective_ppr("ProbeStyle")
        assert integer(effective.ind.left) == 720, name
        blocks = [b for b in extract(package).blocks if b.text.startswith("I62")]
        assert [b.indent_left for b in blocks] == [720], name
    styles = _load(INTEGERS / "cr065-I.docx").main_document_part.style_definitions_part
    probe = next(s for s in styles.contents.style if s.style_id == "ProbeStyle")
    assert probe.p_pr.ind.left == "720.5"  # kept as read: the reader, not the tree, reads it as Word


def test_the_font_size_reads_as_word() -> None:
    """F41's ``w:sz="22.5"`` is 22 half-points, 11 pt, through ``read_run_options`` (``Font``), as in Word's save.

    Word's save states no size (22 is the document default's), so the effective
    reading is what agrees; the direct one is 11 here, as stated.
    """
    from docx4j_py.wml import read_run_options

    readings = []
    for name in ("cr065-F", "cr065-F-word365"):
        paragraph = next(p for p in _load(INTEGERS / f"{name}.docx").body.paragraphs if p.text.startswith("F41"))
        readings.append(paragraph.font.size)
    assert readings == [11.0, 11.0]
    paragraph = next(p for p in _load(INTEGERS / "cr065-F.docx").body.paragraphs if p.text.startswith("F41"))
    run = next(r for r in paragraph.element.content if getattr(r, "r_pr", None) is not None and r.r_pr.sz is not None)
    assert run.r_pr.sz.val == "22.5" and read_run_options(run.r_pr)["size"] == 11.0
