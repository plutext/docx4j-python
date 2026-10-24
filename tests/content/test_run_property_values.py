"""Run properties read as Word 365 reads them: CR-018 Phases A and B, over Word's own saves.

The fixtures are ``tests/fixtures/onoff-fun/`` (``SHA1SUMS``; ``tests/README.md``
has their provenance): the twenty probes of rounds one to four, as docx4j's
``runprops/word365/`` holds them at ``868a1797b``, and the eight of rounds five
and six from the share, and the five of round seven from docx4j at ``5163c2986``, each with
Word 365's save of it (``<name>-word365.docx``).

**The oracle** is docx4j's ``RunPropertiesWord365Test``'s. Word's save of a
probe holds only Word's own spellings (bare for on, a value dropped where Word
reads it as absent), which every reader here reads correctly, so a reader
**agrees with Word** on a marker when its answer on the probe equals its answer
on Word's save. Every probe paragraph after the title is ``label: MARKER
(end)``, and only the marker run carries the property tested.

Phase A reads the eleven on/off properties the probes measure (``b``, ``i``,
``caps``, ``smallCaps``, ``strike``, ``dstrike``, ``outline``, ``shadow``,
``emboss``, ``imprint``, ``vanish``) through every reader of them: the
resolver's toggles, ``Font`` (effective and direct), the text model's runs,
Markdown, the python-docx facade, ``Paragraph.text`` and ``search`` (for
``w:vanish``), and ``format_span``'s write across paragraphs.

Phase B reads the enumerations and colour (``w:u`` and its colour,
``w:highlight``, ``w:vertAlign``, ``w:em``, ``w:color``, ``w:shd``'s fill and
pattern, ``w:bdr``, ``w:background``) through the resolver (docx4j's reading),
``read_run_options`` (effective and direct), the content API's ``Font`` and the
facade. Where Word's own save changes what Word drew (docx4j CR-060 section 13
item 5), Word's answer comes from the PDF and is stated in :data:`WORD_DREW`,
none of which is an on/off reading. The saves are held both ways: the on/off
rows (Phase A) and the enumeration and colour rows (Phase B) of section 4.7.
"""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

import pytest

from docx4j_py.model.values import OFF, is_on, on_off, xsd_boolean

ROOT = Path(__file__).resolve().parents[2]
ONOFF = ROOT / "tests" / "fixtures" / "onoff-fun"

# the probes hold values no rule converts on purpose ("off", "Navy", "bogus"): each is kept and warned about
pytestmark = pytest.mark.filterwarnings("ignore::docx4j_xsdata.exceptions.ConverterWarning")

#: Rounds one to four (docx4j's ``RunPropertiesWord365Test.FILES``), then rounds five and six, then round seven.
PROBES = (
    "onoff-valid",
    "onoff-empty",
    "onoff-space",
    "onoff-TRUE",
    "onoff-yes",
    "props-b-case",
    "props-b-styles",
    "props-toggles-1",
    "props-toggles-2",
    "props-u-valid",
    "props-u-case",
    "props-u-other",
    "props-enum-missing",
    "props-enum-other",
    "colour-names-highlight",
    "colour-names-css",
    "colour-case",
    "colour-hex",
    "colour-table-grey",
    "colour-table-other",
    "props-u-styles",
    "colour-styles",
    "props-enum-styles",
    "colour-elsewhere",
    "props-u-onoff-styles",
    "colour-auto-styles",
    "highlight-char-styles",
    "colour-elsewhere-styles",
    "colour-pattern",
    "colour-background-red",
    "colour-background-navy",
    "colour-background-f00",
    "colour-background-hex",
)

#: The on/off run properties the probes measure: (``w:rPr`` field, facade ``Font`` attribute).
TOGGLES = (
    ("b", "bold"),
    ("i", "italic"),
    ("caps", "all_caps"),
    ("small_caps", "small_caps"),
    ("strike", "strike"),
    ("dstrike", "double_strike"),
    ("outline", "outline"),
    ("shadow", "shadow"),
    ("emboss", "emboss"),
    ("imprint", "imprint"),
    ("vanish", "hidden"),
)

#: Where Word's own save changes what Word drew (docx4j CR-060 section 13 item 5, its ``WORD_DREW``): Word's
#: reading from the PDF. None is an on/off reading, so Phase A's comparisons are the saves'; Phase B reads them.
WORD_DREW: dict[tuple[str, str], dict[str, str]] = {
    ("props-enum-styles", "M6_HIGHLIGHT_GREEN"): {"highlight": "green"},
    ("highlight-char-styles", "H7_CHAR_STYLE_DIRECT_GREEN"): {"highlight": "green"},
    ("colour-elsewhere-styles", "U4_UNDERLINE_AUTO"): {"u-color": "auto"},
}


# ---------------------------------------------------------------------------
# the probes
# ---------------------------------------------------------------------------


class Marker:
    """A probe paragraph's marker run, in the probe and in Word's save."""

    __slots__ = ("index", "p", "probe", "r", "saved_p", "saved_r", "text")

    def __init__(self, probe: str, index: int, text: str, p: Any, r: Any, saved_p: Any, saved_r: Any) -> None:
        self.probe, self.index, self.text = probe, index, text
        self.p, self.r, self.saved_p, self.saved_r = p, r, saved_p, saved_r

    def __repr__(self) -> str:
        return f"{self.probe} {self.text}"


def _runs(p: Any) -> list[Any]:
    from docx4j_py.wml import R

    return [item for item in p.content or () if isinstance(item, R)]


def _text(r: Any) -> str:
    from docx4j_py.traversal import text_of

    return text_of(r)


def _paragraphs(package: Any) -> list[Any]:
    from docx4j_py.wml import P

    return [item for item in package.main_document_part.contents.body.content or () if isinstance(item, P)]


class Probe:
    """A probe and Word's save, loaded, with their markers."""

    def __init__(self, name: str) -> None:
        from docx4j_py.openpackaging.api import load

        self.name = name
        self.path = ONOFF / f"{name}.docx"
        self.saved_path = ONOFF / f"{name}-word365.docx"
        self.package = load(self.path)
        self.saved = load(self.saved_path)
        ps, saved_ps = _paragraphs(self.package), _paragraphs(self.saved)
        assert len(ps) == len(saved_ps), f"{name}: paragraphs in the save"
        self.markers: list[Marker] = []
        for index in range(1, len(ps)):  # the first paragraph is the probe's title
            runs = _runs(ps[index])
            if len(runs) != 3:
                continue
            text = _text(runs[1])
            found = [r for r in _runs(saved_ps[index]) if text in _text(r)]
            assert found, f"{name}: {text} as one run in the save"
            self.markers.append(Marker(name, index, text, ps[index], runs[1], saved_ps[index], found[-1]))
        assert self.markers, f"{name}: markers"


@pytest.fixture(scope="module")
def probes() -> dict[str, Probe]:
    return {name: Probe(name) for name in PROBES}


def test_the_fixtures_are_as_copied() -> None:
    """Every file ``SHA1SUMS`` lists, byte for byte, and nothing else beside it."""
    listed = {}
    for line in (ONOFF / "SHA1SUMS").read_text(encoding="utf-8").splitlines():
        digest, name = line.split("  ", 1)
        listed[name] = digest
    for name, digest in listed.items():
        assert hashlib.sha1((ONOFF / name).read_bytes()).hexdigest() == digest, name
    assert sorted(p.name for p in ONOFF.iterdir() if p.name != "SHA1SUMS") == sorted(listed)
    assert len([n for n in listed if n.endswith(".docx")]) == 2 * len(PROBES) == 66


def test_every_probe_has_markers(probes: dict[str, Probe]) -> None:
    counts = {name: len(probe.markers) for name, probe in probes.items()}
    assert counts["onoff-valid"] == 8
    assert counts["props-b-case"] == 10
    assert counts["props-toggles-1"] == counts["props-toggles-2"] == 26  # five values of five properties, and a control
    assert sum(counts[name] for name in PROBES[:20]) == 227  # rounds one to four: CR-018 section 2's 227
    assert sum(counts[name] for name in PROBES[20:28]) == 149  # rounds five and six
    assert sum(counts[name] for name in PROBES[28:]) == 12  # round seven: P1 to P8, and B1 in each page colour


# ---------------------------------------------------------------------------
# the rule
# ---------------------------------------------------------------------------

#: Rule 1's spellings (CR-018 section 1.1), with Word's reading of each.
SPELLINGS = (
    (None, True),
    (True, True),
    (False, False),
    ("true", True),
    ("on", True),
    ("1", True),
    ("", True),
    (" ", True),
    ("TRUE", True),
    ("FALSE", True),
    ("OFF", True),
    ("False", True),
    ("Off", True),
    (" false", True),
    ("false ", True),
    ("00", True),
    ("no", True),
    ("yes", True),
    ("false", False),
    ("off", False),
    ("0", False),
)


@pytest.mark.parametrize(("value", "word"), SPELLINGS)
def test_on_off_reads_as_word_reads(value: Any, word: bool) -> None:
    assert on_off(value) is word


def test_on_off_defaults_and_members() -> None:
    from docx4j_py.shared_types import StOnOff as SharedTypesStOnOff
    from docx4j_py.w14 import StOnOff

    assert OFF == {"false", "off", "0"}
    assert on_off(None, default=False) is False
    assert on_off(None, default=True) is True
    assert on_off(0) is False and on_off(1) is True
    assert on_off(StOnOff.VALUE_0) is False and on_off(StOnOff.TRUE) is True
    assert on_off(SharedTypesStOnOff.OFF) is False and on_off(SharedTypesStOnOff.ON) is True


def test_is_on() -> None:
    from docx4j_py.wml import el

    assert is_on(None) is False
    assert is_on(el.b()) is True
    assert is_on(el.b(val=False)) is False
    assert is_on(el.b(val="off")) is False
    assert is_on(el.b(val=" false")) is True
    assert is_on(el.b(val="")) is True


@pytest.mark.parametrize(
    ("value", "default", "read"),
    [
        (None, True, True),
        (None, None, None),
        (True, None, True),
        ("1", None, True),
        (" true\t", None, True),
        ("\n0 ", None, False),
        ("false", None, False),
        ("off", True, True),  # not xsd:boolean: the default
        ("yes", False, False),
    ],
)
def test_xsd_boolean(value: Any, default: Any, read: Any) -> None:
    assert xsd_boolean(value, default) is read


# ---------------------------------------------------------------------------
# the readers, against Word's saves
# ---------------------------------------------------------------------------


def _resolver_reading(package: Any, p: Any, r: Any) -> dict[str, bool]:
    resolver = package.get_property_resolver()
    effective = resolver.get_effective_rpr(r.r_pr, p.p_pr)
    return {field: is_on(getattr(effective, field, None)) for field, _ in TOGGLES}


def _compare(readings: list[tuple[Marker, Any, Any]]) -> list[str]:
    return [f"{marker}: probe {probe!r}, Word {word!r}" for marker, probe, word in readings if probe != word]


def test_the_resolver_reads_as_word(probes: dict[str, Probe]) -> None:
    """The resolver's effective toggles (``is_true`` and ``StyleUtil.toggle``), all eleven, every marker."""
    readings = [
        (m, _resolver_reading(pr.package, m.p, m.r), _resolver_reading(pr.saved, m.saved_p, m.saved_r))
        for pr in probes.values()
        for m in pr.markers
    ]
    assert not _compare(readings), "\n".join(_compare(readings))


def _font(package: Any, p: Any, r: Any, *, direct: bool) -> dict[str, Any]:
    from docx4j_py.wml import read_run_options

    if direct:
        options = read_run_options(r.r_pr)
    else:
        options = read_run_options(package.get_property_resolver().get_effective_rpr(r.r_pr, p.p_pr))
    return {k: options[k] for k in ("bold", "italic", "strike_through", "double_strike_through")}


@pytest.mark.parametrize("direct", [False, True], ids=["effective", "direct"])
def test_font_reads_as_word(probes: dict[str, Probe], direct: bool) -> None:
    """``read_run_options``, the reader under ``Font`` (``bold``, ``italic``, ``strike_through``, ``double_strike_through``)."""
    readings = [
        (m, _font(pr.package, m.p, m.r, direct=direct), _font(pr.saved, m.saved_p, m.saved_r, direct=direct))
        for pr in probes.values()
        for m in pr.markers
    ]
    assert not _compare(readings), "\n".join(_compare(readings))


def test_the_content_api_font_reads_as_word(probes: dict[str, Probe]) -> None:
    """``Paragraph`` ranges' ``Font`` over each marker run (the first run in scope's effective formatting)."""

    def reading(package: Any, index: int, text: str) -> Any:
        found = package.body.paragraphs[index].search(text, match_case=True)
        if not found:
            return "hidden"
        font = found[0].font
        return (font.bold, font.italic, font.strike_through, font.double_strike_through)

    readings = [
        (m, reading(pr.package, m.index, m.text), reading(pr.saved, m.index, m.text))
        for pr in probes.values()
        for m in pr.markers
    ]
    assert not _compare(readings), "\n".join(_compare(readings))


def _text_runs(package: Any) -> list[Any]:
    from docx4j_py.model.text import extract

    out = []
    for block in extract(package).blocks:
        out.extend(getattr(block, "runs", None) or ())
    return out


def test_the_text_model_reads_as_word(probes: dict[str, Probe]) -> None:
    """The text model's runs: bold, italic and struck through, and hidden runs left out."""

    def reading(runs: list[Any], text: str) -> Any:
        found = [run for run in runs if text in (run.text or "")]
        if not found:
            return "hidden"
        return (found[0].bold, found[0].italic, found[0].strike)

    readings = []
    for pr in probes.values():
        runs, saved = _text_runs(pr.package), _text_runs(pr.saved)
        readings.extend((m, reading(runs, m.text), reading(saved, m.text)) for m in pr.markers)
    assert not _compare(readings), "\n".join(_compare(readings))


def test_markdown_reads_as_word(probes: dict[str, Probe]) -> None:
    """Each marker's line of Markdown is the same for the probe and for Word's save."""
    from docx4j_py.model.text import extract, markdown

    def lines(package: Any) -> dict[str, str]:
        out = {}
        for line in markdown(extract(package), package).splitlines():
            label = line.split(":", 1)[0]
            out.setdefault(label, line)
        return out

    readings = []
    for pr in probes.values():
        probe, word = lines(pr.package), lines(pr.saved)
        for m in pr.markers:
            label = _text(_runs(m.p)[0]).split(":", 1)[0]
            readings.append((m, probe.get(label), word.get(label)))
    assert not _compare(readings), "\n".join(_compare(readings))


def test_the_facade_reads_as_word(probes: dict[str, Probe]) -> None:
    """python-docx's ``Font`` over the marker run: all eleven, python-docx's tri-state shape (CR-018 Q9)."""
    from docx4j_py.docx import Document

    def reading(document: Any, index: int, text: str) -> dict[str, Any]:
        run = next(run for run in document.paragraphs[index].runs if text in run.text)
        # tri-state: None (not stated) and False (stated off) are one reading, as Word's save drops an off value
        return {attribute: bool(getattr(run.font, attribute)) for _, attribute in TOGGLES}

    readings = []
    for pr in probes.values():
        probe, word = Document(str(pr.path)), Document(str(pr.saved_path))
        readings.extend((m, reading(probe, m.index, m.text), reading(word, m.index, m.text)) for m in pr.markers)
    assert not _compare(readings), "\n".join(_compare(readings))


def test_hidden_text_reads_as_word(probes: dict[str, Probe]) -> None:
    """``w:vanish`` (T10): ``Paragraph.text`` and ``search`` read a ``w:val="off"`` run's text, as Word shows it."""
    pr = probes["props-toggles-2"]
    for m in pr.markers:
        if not m.text.startswith("sample t10"):
            continue
        assert pr.package.body.paragraphs[m.index].text == pr.saved.body.paragraphs[m.index].text, m
        assert bool(pr.package.body.search(m.text)) == bool(pr.saved.body.search(m.text)), m
    off = next(m for m in pr.markers if m.text == "sample t10 vanish off")
    assert pr.package.body.paragraphs[off.index].text == 'T10 w:vanish w:val="off": sample t10 vanish off (end)'
    assert len(pr.package.body.search("t10 vanish off")) == 1


def test_bold_across_paragraphs_writes_over_off(probes: dict[str, Probe]) -> None:
    """``range.font.bold = True`` across V4 to V6 writes ``w:b`` over V5's ``"off"`` (``format_span``'s ``_reads``)."""
    from docx4j_py.openpackaging.api import load

    package = load(ONOFF / "onoff-valid.docx")
    body = package.body
    span = body.search("V4_ON")[0].expand_to(body.search("V6_ONE")[0])
    span.font.bold = True
    v5 = next(r for r in _runs(body.paragraphs[5].element) if _text(r) == "V5_OFF")
    assert is_on(v5.r_pr.b), v5.r_pr.b
    assert v5.r_pr.b.val != "off"
    assert body.search("V5_OFF")[0].font.bold is True


# ---------------------------------------------------------------------------
# style levels
# ---------------------------------------------------------------------------


def test_a_style_level_off_is_a_stated_false() -> None:
    """``StyleUtil.toggle``: a character style stating ``"off"`` over a bold paragraph style leaves it bold.

    No probe has the case; it follows from rule 1 and the toggle rule CR-002
    section 15 settled (a stated false at a style level is a term of the XOR
    that inverts nothing).
    """
    from docx4j_py.model.style_util import toggle
    from docx4j_py.wml import el

    assert is_on(toggle(el.b(val="off"), el.b(), None)) is True
    assert is_on(toggle(el.b(val=False), el.b(), None)) is True
    assert is_on(toggle(el.b(), el.b(), None)) is False  # two bold levels cancel
    assert is_on(toggle(el.b(val=""), el.b(), None)) is False  # "" is a stated true
    assert is_on(toggle(el.b(val="off"), None, None)) is False  # its own false stands


# ---------------------------------------------------------------------------
# what a save writes (CR-018 section 4.7, Q5, Q17)
# ---------------------------------------------------------------------------


def _parts(path: Path) -> list[str]:
    import zipfile

    return sorted(n for n in zipfile.ZipFile(path).namelist() if n.startswith("word/") and n.endswith(".xml"))


def _saved(path: Path, **save: Any) -> tuple[Any, dict[str, bytes], dict[str, bytes]]:
    """The probe loaded, every WordprocessingML part read (so re-marshalled), and saved."""
    import io
    import zipfile

    from docx4j_py.openpackaging.api import load

    package = load(path)
    for name in _parts(path):
        part = package.get_part("/" + name)
        if part is not None and hasattr(part, "contents"):
            _ = part.contents
    data = package.save(**save)
    source = {n: zipfile.ZipFile(path).read(n) for n in _parts(path)}
    written = {n: zipfile.ZipFile(io.BytesIO(data)).read(n) for n in _parts(path)}
    return package, source, written


@pytest.mark.parametrize("name", PROBES)
def test_a_save_that_keeps_the_spelling_changes_nothing(name: str) -> None:
    """``save(normalise=False)``: every part canonically identical to its source, ``boolean-spelling`` aside.

    CR-018 section 2.7's second row: the exact parse keeps ``" false"`` and
    ``"false "``, which the trimming converter wrote ``0``.
    """
    import canon

    package, source, written = _saved(ONOFF / f"{name}.docx", normalise=False)
    for part, data in source.items():
        report = canon.compare(data, written[part])
        assert report.identical, (part, report.differences, report.diffs[:5])
        assert "word-normal-form" not in report.counts
    assert package.last_save.normalise is False and package.last_save.normalised == []


@pytest.mark.parametrize("name", PROBES)
def test_a_normalising_save_moves_only_the_kept_values(name: str) -> None:
    """``save()``, normalising: the same normal form written into the source, nothing else differs."""
    import canon

    package, source, written = _saved(ONOFF / f"{name}.docx")
    rewrites = 0
    for part, data in source.items():
        report = canon.compare(data, written[part], normalise=True)
        assert report.identical, (part, report.differences, report.diffs[:5])
        rewrites += report.counts.get("word-normal-form", 0)
    assert rewrites == len(package.last_save.normalised)
    # the on/off rows are values the parse kept (Phase A); the enumeration and colour rows are found in
    # the render (Phase B), so a lower-case hex or a bare <w:u/>, which bind, are among them too
    kept = {k.text for _p, k in package.kept}
    for record in package.last_save.normalised:
        if _on_off_row(record):
            assert record.before in kept, record


#: The on/off values each probe's normalising save rewrites, and how (CR-018 section 4.7's on/off rows).
NORMALISED = {
    "onoff-valid": {"on": None, "off": "0"},
    "onoff-empty": {"": None},
    "onoff-space": {" ": None},
    "onoff-TRUE": {"TRUE": None},
    "onoff-yes": {"yes": None},
    "props-b-case": {"FALSE": None, "OFF": None, "False": None, "Off": None, " false": None, "false ": None,
                     "00": None, "no": None},
    "props-b-styles": {"off": "0"},
}


def _on_off_row(record: Any) -> bool:
    """A Phase A row (an on/off value), as against Phase B's enumeration and colour rows.

    Phase B's are an element dropped (a path with no attribute) or a ``w:val``,
    ``w:color`` or ``w:fill`` of one of :data:`PHASE_B_ELEMENTS`.
    """
    last = record.path.rsplit("/", 1)[-1]
    if not last.startswith("@"):
        return False
    element = record.path.rsplit("/", 2)[-2].split("[", 1)[0]
    return not (element in PHASE_B_ELEMENTS and last in ("@w:val", "@w:color", "@w:fill"))


#: The elements Phase B's rows rewrite (CR-018 section 4.7).
PHASE_B_ELEMENTS = frozenset(
    f"w:{name}"
    for name in ("color", "u", "highlight", "vertAlign", "em", "shd", "bdr", "background", "top", "left", "bottom",
                 "right", "start", "end", "between", "bar", "insideH", "insideV", "tl2br", "tr2bl")
)


@pytest.mark.parametrize("name", sorted(NORMALISED))
def test_the_on_off_rows(name: str) -> None:
    package, _source, _written = _saved(ONOFF / f"{name}.docx")
    on_off_rows = [record for record in package.last_save.normalised if _on_off_row(record)]
    found = {record.before: record.after for record in on_off_rows}
    assert found == NORMALISED[name]
    assert all(record.part == "/word/document.xml" for record in on_off_rows)
    assert all(record.path.endswith("/w:rPr/w:b/@w:val") for record in on_off_rows)


def test_the_toggles_rows() -> None:
    """Rounds one and two's other toggles: ``"off"`` written ``0``; ``""`` and ``"FALSE"`` bare."""
    for name in ("props-toggles-1", "props-toggles-2"):
        package, _source, _written = _saved(ONOFF / f"{name}.docx")
        by_value: dict[str, set[Any]] = {}
        for record in package.last_save.normalised:
            by_value.setdefault(record.before, set()).add(record.after)
        assert by_value == {"off": {"0"}, "": {None}, "FALSE": {None}}, (name, by_value)
        assert len(package.last_save.normalised) == 15


@pytest.mark.parametrize("name", PROBES)
def test_word_reads_a_normalised_save_as_it_read_the_source(name: str, tmp_path: Path) -> None:
    """The oracle again, over this package's normalising save of each probe: every marker reads as in Word's save."""
    path = tmp_path / f"{name}.docx"
    _package, _source, _written = _saved(ONOFF / f"{name}.docx")
    path.write_bytes(_package.save())
    from docx4j_py.openpackaging.api import load

    ours = load(path)
    word = Probe(name)
    ps = _paragraphs(ours)
    for m in word.markers:
        r = next(r for r in _runs(ps[m.index]) if m.text in _text(r))
        assert _resolver_reading(ours, ps[m.index], r) == _resolver_reading(word.saved, m.saved_p, m.saved_r), m


def test_the_tree_keeps_what_was_read_and_the_option() -> None:
    """The normal form is the save's: the tree keeps the source's values, and the load option turns it off."""
    import io
    import zipfile

    from docx4j_py.openpackaging.api import load
    from docx4j_py.openpackaging.load import LoadOptions

    package = load(ONOFF / "onoff-valid.docx")
    v4 = _runs(_paragraphs(package)[4])[1]
    assert v4.r_pr.b.val == "on"
    first = package.save()
    assert v4.r_pr.b.val == "on"  # taken back out after the save
    assert len(package.last_save.normalised) == 2
    second = package.save(normalise=False)
    # (the labels' text holds the spellings too, so the element is what is looked for)
    assert b'<w:b w:val="on"/>' not in zipfile.ZipFile(io.BytesIO(first)).read("word/document.xml")
    assert b'<w:b w:val="on"/>' in zipfile.ZipFile(io.BytesIO(second)).read("word/document.xml")

    options = LoadOptions(normalise_on_save=False)
    kept = load(ONOFF / "onoff-valid.docx", options=options)
    _ = kept.main_document_part.contents
    assert b'<w:b w:val="off"/>' in zipfile.ZipFile(io.BytesIO(kept.save())).read("word/document.xml")
    assert kept.last_save.normalise is False
    normalised = zipfile.ZipFile(io.BytesIO(kept.save(normalise=True))).read("word/document.xml")
    assert b'<w:b w:val="off"/>' not in normalised and b'<w:b w:val="0"/>' in normalised


def test_a_copied_value_is_normalised_too() -> None:
    """A kept value an edit copied (here a run copied into a new paragraph) is written in the normal form as well."""
    import io
    import zipfile

    from docx4j_py.child import deep_copy
    from docx4j_py.openpackaging.api import load
    from docx4j_py.wml import el

    package = load(ONOFF / "onoff-valid.docx")
    body = package.main_document_part.contents.body
    v5 = _runs(_paragraphs(package)[5])[1]
    body.content.insert(0, el.p(content=[deep_copy(v5)]))
    xml = zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml").decode()
    assert xml.count('<w:b w:val="off"/>') == 0 and xml.count('<w:b w:val="0"/>') == 4  # V3, V5, its copy, V7
    assert len([r for r in package.last_save.normalised if r.before == "off"]) == 2


def test_a_value_set_in_code_keeps_today_s_form() -> None:
    import io
    import zipfile

    from docx4j_py.openpackaging.api import create_package
    from docx4j_py.wml import el, p, r

    package = create_package()
    body = package.main_document_part.contents.body
    body.content.append(p(r("x", bold=True)))
    run = body.content[-1].content[0]
    run.r_pr.i = el.i(val=False)
    xml = zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml").decode()
    assert "<w:b/>" in xml and '<w:i w:val="0"/>' in xml
    assert package.last_save.normalised == []


def test_describe_counts_the_kept_values() -> None:
    from docx4j_py.openpackaging.api import load

    package = load(ONOFF / "props-b-case.docx")
    _ = package.main_document_part.contents
    assert package.describe().kept == {"exact": 2, "invalid": 6}
    assert package.describe().to_dict()["kept"] == {"exact": 2, "invalid": 6}


@pytest.mark.parametrize(
    ("value", "absent_is_on", "written"),
    [("on", True, None), ("", True, None), (" false", True, None), ("off", True, False),
     ("on", False, True), ("off", False, False), ("TRUE", False, True)],
)
def test_on_off_normal_form(value: str, absent_is_on: bool, written: Any) -> None:
    from docx4j_py.model.values import on_off_normal_form

    assert on_off_normal_form(value, absent_is_on) is written


# ===========================================================================
# Phase B: the enumerations and colour (CR-018 sections 4.2, 4.3, 4.7)
# ===========================================================================

#: ``w:highlight``'s values Word draws, as ``read_run_options`` reports them (docx4j's ``Highlight.getHexVal``).
_HIGHLIGHT_HEX = {
    "black": "#000000", "blue": "#0000FF", "cyan": "#00FFFF", "green": "#008000", "magenta": "#FF00FF",
    "red": "#FF0000", "yellow": "#FFFF00", "white": "#FFFFFF", "darkBlue": "#00008B", "darkCyan": "#008B8B",
    "darkGreen": "#006400", "darkMagenta": "#8B008B", "darkRed": "#8B0000", "darkYellow": "#FFD700",
    "darkGray": "#A9A9A9", "lightGray": "#D3D3D3",
}


def _drawn_colour(value: Any, automatic: str) -> str:
    from docx4j_py.model.values import AUTO, colour

    read = colour(value)
    return automatic if read in (None, AUTO) else read


def _drawn(package: Any, p: Any, r: Any) -> dict[str, str]:
    """docx4j's ``RunPropertiesWord365Test.reading``, its enumeration and colour half: what Word draws.

    The effective ``w:rPr`` through the resolver, each value read by
    :mod:`docx4j_py.model.values`: the underline and its colour, the highlight,
    the vertical alignment, the emphasis mark, the text colour, the shading's
    fill and pattern colour, and the border's colour. A stated automatic is
    automatic, and an absent shading's pattern is too.
    """
    from docx4j_py.model.values import (
        emphasis_mark,
        highlight,
        underline,
        vertical_alignment,
    )

    effective = package.get_property_resolver().get_effective_rpr(r.r_pr, p.p_pr)
    out: dict[str, str] = {}
    u = underline(getattr(effective.u, "val", None))
    out["u"] = "none" if u is None or u.value == "none" else u.value
    h = highlight(getattr(effective.highlight, "val", None))
    out["highlight"] = "none" if h is None or h.value == "none" else h.value
    va = vertical_alignment(getattr(effective.vert_align, "val", None))
    out["vertAlign"] = "none" if va is None else va.value
    em = emphasis_mark(getattr(effective.em, "val", None))
    out["em"] = "none" if em is None or em.value == "none" else em.value
    out["color"] = _drawn_colour(getattr(effective.color, "val", None), "auto")
    out["u-color"] = "-" if out["u"] == "none" else _drawn_colour(getattr(effective.u, "color", None), "auto")
    shd = effective.shd
    out["fill"] = _drawn_colour(getattr(shd, "fill", None), "none")
    out["pattern"] = _drawn_colour(getattr(shd, "color", None), "auto")
    bdr = effective.bdr
    kind = getattr(getattr(bdr, "val", None), "value", None)
    out["border"] = _drawn_colour(bdr.color, "000000") if kind not in (None, "none", "nil") else "none"
    return out


def _word_drew(marker: Any, reading: dict[str, Any], reader: str) -> dict[str, Any]:
    """Word's reading of a marker: its save's, with :data:`WORD_DREW`'s answer from the PDF where the save changes it."""
    over = WORD_DREW.get((marker.probe, marker.text))
    if not over:
        return reading
    reading = dict(reading)
    if "highlight" in over:
        drawn = over["highlight"]
        if reader == "drawn":
            reading["highlight"] = drawn
        elif reader == "options":
            reading["highlight_color"] = _HIGHLIGHT_HEX[drawn]
        elif reader == "facade":
            from docx4j_py.docx.enums import WD_COLOR_INDEX

            reading["highlight_color"] = next(m for m in WD_COLOR_INDEX if m.xml_value == drawn)
    if "u-color" in over and reader == "drawn":
        reading["u-color"] = over["u-color"]
    return reading


def test_the_markers_of_round_seven(probes: dict[str, Probe]) -> None:
    assert [m.text for m in probes["colour-pattern"].markers] == [
        "P1_PATTERN_RED", "P2_PATTERN_NAVY", "P3_PATTERN_F00", "P4_PATTERN_HEX", "P5_PATTERN_AUTO",
        "P6_STYLE", "P7_STYLE_F00", "P8_STYLE_NAVY",
    ]
    for name in PROBES[29:]:
        assert [m.text for m in probes[name].markers] == ["B1_PAGE"]


def test_the_resolver_draws_as_word(probes: dict[str, Probe]) -> None:
    """The resolver's effective enumerations and colours (docx4j's oracle), every marker of every round."""
    readings = [
        (m, _drawn(pr.package, m.p, m.r), _word_drew(m, _drawn(pr.saved, m.saved_p, m.saved_r), "drawn"))
        for pr in probes.values()
        for m in pr.markers
    ]
    assert len(readings) == 388
    assert not _compare(readings), "\n".join(_compare(readings))


_OPTIONS = ("underline", "highlight_color", "color", "subscript", "superscript")


def _options(r_pr: Any) -> dict[str, Any]:
    from docx4j_py.wml import read_run_options

    read = read_run_options(r_pr)
    out = {k: read[k] for k in _OPTIONS}
    if out["color"] == "auto":  # a stated automatic is drawn as no colour: Word's save drops one nothing needs
        out["color"] = ""
    return out


def _excused(marker: Any, probe: dict[str, Any], word: dict[str, Any], drawn_agree: bool) -> dict[str, Any]:
    """Word's direct reading, where its save left out a direct value its style already gives.

    Word's save writes a run's direct formatting relative to its styles: over
    ``CharUnderlineRed`` a direct ``w:u w:val="single" w:color="Navy"`` is saved
    ``<w:u w:color="000080"/>``. Where the save states nothing for a property
    and the effective readings agree (:func:`_drawn`), the probe's direct value
    is the one Word read, and the save's silence is not a reading.
    """
    r_pr = marker.saved_r.r_pr
    stated = {
        "underline": getattr(getattr(r_pr, "u", None), "val", None),
        "highlight_color": getattr(r_pr, "highlight", None),
        "color": getattr(getattr(r_pr, "color", None), "val", None),
        "vert_align": getattr(r_pr, "vert_align", None),
    }
    stated["rgb"] = stated["type"] = stated["color"]
    stated["subscript"] = stated["superscript"] = stated["vert_align"]
    word = dict(word)
    for key, value in probe.items():
        if value != word.get(key) and drawn_agree and stated[key] is None:
            word[key] = value
    return word


@pytest.mark.parametrize("direct", [False, True], ids=["effective", "direct"])
def test_read_run_options_reads_the_enumerations_and_colour_as_word(probes: dict[str, Probe], direct: bool) -> None:
    """``read_run_options``, the reader under ``Font``: ``underline``, ``highlight_color``, ``color``, sub- and superscript."""
    readings = []
    for pr in probes.values():
        resolver, saved_resolver = pr.package.get_property_resolver(), pr.saved.get_property_resolver()
        for m in pr.markers:
            if direct:
                probe, word = _options(m.r.r_pr), _options(m.saved_r.r_pr)
                agree = _drawn(pr.package, m.p, m.r) == _word_drew(m, _drawn(pr.saved, m.saved_p, m.saved_r), "drawn")
                word = _excused(m, probe, _word_drew(m, word, "options"), agree)
            else:
                probe = _options(resolver.get_effective_rpr(m.r.r_pr, m.p.p_pr))
                word = _word_drew(m, _options(saved_resolver.get_effective_rpr(m.saved_r.r_pr, m.saved_p.p_pr)), "options")
            readings.append((m, probe, word))
    assert not _compare(readings), "\n".join(_compare(readings))


def test_the_content_api_font_reads_the_enumerations_and_colour_as_word(probes: dict[str, Probe]) -> None:
    """``Font.underline``, ``highlight_color``, ``color``, ``subscript`` and ``superscript`` over each marker run."""

    def reading(package: Any, index: int, text: str) -> Any:
        found = package.body.paragraphs[index].search(text, match_case=True)
        if not found:
            return "hidden"
        font = found[0].font
        out = {k: getattr(font, k) for k in _OPTIONS}
        if out["color"] == "auto":
            out["color"] = ""
        return out

    readings = []
    for pr in probes.values():
        for m in pr.markers:
            probe, word = reading(pr.package, m.index, m.text), reading(pr.saved, m.index, m.text)
            if isinstance(word, dict):
                word = _word_drew(m, word, "options")
            readings.append((m, probe, word))
    assert not _compare(readings), "\n".join(_compare(readings))


def test_the_facade_reads_the_enumerations_and_colour_as_word(probes: dict[str, Probe]) -> None:
    """python-docx's ``Font``, Word's reading in python-docx's shape (CR-018 Q9): direct, tri-state.

    ``underline`` stated ``none`` and absent are one reading, as ``AUTO`` and no
    ``w:color`` are, since Word's save drops either where nothing beneath needs it.
    """
    from docx4j_py.docx import Document

    def reading(document: Any, index: int, text: str) -> dict[str, Any]:
        font = next(run for run in document.paragraphs[index].runs if text in run.text).font
        kind = font.color.type
        return {
            "underline": None if font.underline is False else font.underline,
            "highlight_color": font.highlight_color,
            "superscript": font.superscript,
            "subscript": font.subscript,
            "rgb": None if font.color.rgb is None else str(font.color.rgb),
            "type": None if kind is not None and kind.name == "AUTO" else kind,
        }

    readings = []
    for pr in probes.values():
        probe, word = Document(str(pr.path)), Document(str(pr.saved_path))
        for m in pr.markers:
            ours, theirs = reading(probe, m.index, m.text), _word_drew(m, reading(word, m.index, m.text), "facade")
            agree = _drawn(pr.package, m.p, m.r) == _word_drew(m, _drawn(pr.saved, m.saved_p, m.saved_r), "drawn")
            readings.append((m, ours, _excused(m, ours, theirs, agree)))
    assert not _compare(readings), "\n".join(_compare(readings))


def _page_colour(package: Any) -> str:
    background = package.main_document_part.contents.background
    return "none" if background is None else _drawn_colour(background.color, "none")


def test_the_page_colour_reads_as_word(probes: dict[str, Probe]) -> None:
    """``w:background/@w:color`` by :func:`colour` (round seven): a name read, a value Word rejects no page colour."""
    found = {name: (_page_colour(probes[name].package), _page_colour(probes[name].saved)) for name in PROBES[29:]}
    assert found == {
        "colour-background-red": ("FF0000", "FF0000"),
        "colour-background-navy": ("000080", "000080"),
        "colour-background-f00": ("none", "none"),  # Word's save drops the element
        "colour-background-hex": ("FF0000", "FF0000"),
    }


def test_word_s_own_save_changes_what_it_drew(probes: dict[str, Probe]) -> None:
    """:data:`WORD_DREW`: the probe reads as Word drew it (the PDFs), where Word's save does not (docx4j's three)."""
    def marker(name: str, text: str) -> Any:
        return next(m for m in probes[name].markers if m.text == text)

    for name, text in (("props-enum-styles", "M6_HIGHLIGHT_GREEN"), ("highlight-char-styles", "H7_CHAR_STYLE_DIRECT_GREEN")):
        m, pr = marker(name, text), probes[name]
        assert _drawn(pr.package, m.p, m.r)["highlight"] == "green", m
        assert _drawn(pr.saved, m.saved_p, m.saved_r)["highlight"] == "none", m  # the save dropped the green
    m, pr = marker("colour-elsewhere-styles", "U4_UNDERLINE_AUTO"), probes["colour-elsewhere-styles"]
    assert _drawn(pr.package, m.p, m.r)["u-color"] == "auto"  # drawn black over the style's red
    assert _drawn(pr.saved, m.saved_p, m.saved_r)["u-color"] == "FF0000"  # the save dropped the w:u


def test_every_colour_marker_reads_as_word_saved_it(probes: dict[str, Probe]) -> None:
    """Rounds two to four's 90 colour markers: ``colour()`` of the probe's ``w:color`` is Word's saved hex, or absent."""
    from docx4j_py.model.values import AUTO, colour

    count = 0
    for name in PROBES[:20]:
        for m in probes[name].markers:
            stated = getattr(getattr(m.r.r_pr, "color", None), "val", None)
            if getattr(m.r.r_pr, "color", None) is None:
                continue
            count += 1
            saved = getattr(getattr(m.saved_r.r_pr, "color", None), "val", None)
            read = colour(stated)
            if read is None or read == AUTO:
                assert saved in (None, colour("auto")) or colour(saved) == AUTO, m
            else:
                assert colour(saved) == read, (m, stated, saved)
    assert count == 90 + 1  # CR-018 section 2.6's 90, and props-enum-missing's M4, a bare <w:color/>


# ---------------------------------------------------------------------------
# the rules as tables
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("value", "read"),
    [
        ("FF0000", "FF0000"), ("ff0000", "FF0000"), ("Ff0000", "FF0000"), ("#FF0000", "FF0000"),
        (" FF0000", "FF0000"), ("FF0000 ", "FF0000"), (" #ff0000 ", "FF0000"),
        ("red", "FF0000"), ("RED", "FF0000"), ("DarkRed", "8B0000"), ("lightgrey", "D3D3D3"), ("navy", "000080"),
        ("auto", "auto"),
        ("Auto", None), ("AUTO", None), (" auto", None), ("auto ", None), ("#auto", None),
        ("F00", None), ("0xFF0000", None), ("FF000", None), ("FF00000", None), ("GG0000", None), ("", None),
        ("grey", None), ("lightgray", None), ("darkgrey", None), ("rebeccapurple", None), ("transparent", None),
        ("windowtext", None), ("dark red", None), ("darkYellow", None), ("FF 00 00", None),
        (None, None), (b"\xff\x00\x00", "FF0000"), (b"", None), (b"\xff\x00", None),
    ],
)
def test_colour(value: Any, read: Any) -> None:
    from docx4j_py.model.values import colour

    assert colour(value) == read


def test_colour_over_lexical_bytes_and_members() -> None:
    """A parsed hexBinary keeps its source's spelling (the fork's ``lexical_types``), and that is what is read."""
    from docx4j_xsdata.models.datatype import XmlLexicalBytes

    from docx4j_py.model.values import AUTO, colour, states_colour
    from docx4j_py.wml import STHexColorAuto

    assert colour(XmlLexicalBytes(b"\xff\x00\x00", "ff0000")) == "FF0000"
    assert colour(XmlLexicalBytes(b"\xff\x00\x00", " FF0000")) == "FF0000"
    assert colour(XmlLexicalBytes(b"\xff\x00\x00", "FF 00 00")) is None  # bytes.fromhex's spaces are not Word's
    assert colour(STHexColorAuto.AUTO) == AUTO == "auto"
    assert states_colour("auto") and states_colour("red") and not states_colour("Auto") and not states_colour(None)


def _docx4j_table(commit: str, path: str, name: str) -> dict[str, str]:
    import re
    import subprocess

    docx4j = ROOT.parent / "docx4j"
    if not (docx4j / ".git").exists():
        pytest.skip("../docx4j is not checked out")
    try:
        source = subprocess.run(
            ["git", "-C", str(docx4j), "show", f"{commit}:{path}"], check=True, capture_output=True, text=True
        ).stdout
    except subprocess.CalledProcessError:
        pytest.skip(f"../docx4j has no {path} at {commit}")
    body = re.search(rf"Map<String, String> {name} = table\((.*?)\);", source, re.DOTALL)
    assert body is not None
    items = re.findall(r'"([^"]+)"', body.group(1))
    return dict(zip(items[::2], items[1::2], strict=True))


def test_the_names_are_docx4j_s() -> None:
    """:data:`NETSCAPE_COLOURS` is docx4j's ``ColorTables.HTML`` at ``868a1797b``, and ``WordColorValues.HTML`` since CR-061."""
    from docx4j_py.model.values import NETSCAPE_COLOURS

    at_cr060 = _docx4j_table(
        "868a1797b", "docx4j-core/src/main/java/org/docx4j/model/color/ColorTables.java", "HTML"
    )
    at_cr061 = _docx4j_table(
        "5163c2986", "docx4j-generated-objects/src/main/java/org/docx4j/wml/WordColorValues.java", "HTML"
    )
    assert NETSCAPE_COLOURS == at_cr060 == at_cr061
    assert len(NETSCAPE_COLOURS) == 140
    css3_only = {"grey", "darkgrey", "dimgrey", "slategrey", "lightslategrey", "darkslategrey", "lightgray"}
    assert not css3_only & set(NETSCAPE_COLOURS)  # round four: Word dropped all seven
    assert all(name == name.lower() for name in NETSCAPE_COLOURS)


def test_the_measured_names(probes: dict[str, Probe]) -> None:
    """Every name rounds two to four gave Word (CR-018 section 1.1): 41 written as their hex at 48 markers, 18 dropped."""
    from docx4j_py.model.values import colour

    accepted: set[str] = set()
    dropped: set[str] = set()
    markers = 0
    for name in ("props-enum-other", "colour-names-highlight", "colour-names-css", "colour-case", "colour-table-grey",
                 "colour-table-other"):
        for m in probes[name].markers:
            value = getattr(getattr(m.r.r_pr, "color", None), "val", None)
            if not isinstance(value, str) or value.strip().lstrip("#").lower() in ("auto", "") or colour(value) is not None and (
                value.strip().lstrip("#").upper() == colour(value)
            ):
                continue  # a hex spelling, or auto: not a name
            saved = getattr(getattr(m.saved_r.r_pr, "color", None), "val", None)
            read = colour(value)
            if read is None:
                assert saved is None, (m, value, saved)
                dropped.add(value.lower())
            else:
                markers += 1
                assert colour(saved) == read, (m, value, saved)
                accepted.add(value.lower())
    assert (len(accepted), markers, len(dropped)) == (41, 48, 18), (sorted(accepted), sorted(dropped))


@pytest.mark.parametrize(
    ("value", "read"),
    [
        ("single", "single"), ("none", "none"), ("dashDotDotHeavy", "dashDotDotHeavy"),
        ("SINGLE", None), ("Single", None), ("NONE", None), ("", None), ("bogus", None), (" single", None),
        ("true", "none"), ("false", "none"), ("on", "none"), ("off", "none"), ("1", "none"), ("0", "none"),
        ("True", None), ("TRUE", None), ("yes", None), ("no", None), (None, None),
    ],
)
def test_underline(value: Any, read: Any) -> None:
    from docx4j_py.model.values import underline

    member = underline(value)
    assert (None if member is None else member.value) == read


def test_in_enumeration() -> None:
    from docx4j_py.model.values import (
        emphasis_mark,
        highlight,
        in_enumeration,
        vertical_alignment,
    )
    from docx4j_py.wml import HighlightVal, STEm, STVerticalAlignRun

    assert highlight("yellow") is HighlightVal.YELLOW and highlight(HighlightVal.RED) is HighlightVal.RED
    assert highlight("YELLOW") is None and highlight("bogus") is None and highlight(" yellow") is None
    assert highlight(None) is None and highlight("") is None and highlight("darkyellow") is None
    assert vertical_alignment("superscript") is STVerticalAlignRun.SUPERSCRIPT
    assert vertical_alignment("Superscript") is None and vertical_alignment("bogus") is None
    assert emphasis_mark("dot") is STEm.DOT and emphasis_mark("DOT") is None
    assert in_enumeration(True, STEm) is None and in_enumeration(1, STEm) is None
    assert in_enumeration("x", STEm, {"x": "dot"}) is STEm.DOT


# ---------------------------------------------------------------------------
# over a style (rounds five and six)
# ---------------------------------------------------------------------------


def _styled_package(styles_xml: str, runs_xml: list[str]) -> Any:
    """A package with these styles added and one paragraph per run."""
    from docx4j_py.fragments import wml
    from docx4j_py.openpackaging.api import create_package

    package = create_package()
    styles = package.main_document_part.style_definitions_part.contents
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    for style in styles_xml.replace("<w:style ", f"<w:style {w} ").split("</w:style>")[:-1]:
        styles.style.append(wml(style + "</w:style>"))
    package.get_property_resolver().refresh()
    body = package.main_document_part.contents.body
    for run in runs_xml:
        body.content.append(wml(run.replace("<w:p>", f"<w:p {w}>", 1)))
    return package


def test_an_invalid_value_over_a_style_leaves_the_style_s() -> None:
    """docx4j's ``anInvalidValueOverAStyleLeavesTheStyles``: highlight, vertAlign, em, and a character style's highlight."""
    styles = (
        '<w:style w:type="character" w:customStyle="1" w:styleId="Marked"><w:name w:val="Marked"/>'
        '<w:rPr><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/><w:em w:val="dot"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="ParaMarked"><w:name w:val="ParaMarked"/>'
        '<w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style>'
    )
    directs = (
        '<w:highlight w:val="YELLOW"/><w:vertAlign w:val="Superscript"/><w:em w:val="DOT"/>',
        '<w:highlight w:val="bogus"/><w:vertAlign w:val="bogus"/><w:em w:val="bogus"/>',
        "<w:highlight/><w:vertAlign/><w:em/>",
    )
    runs = []
    for direct in directs:
        runs.append(f'<w:p><w:pPr><w:pStyle w:val="ParaMarked"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Marked"/>{direct}'
                    "</w:rPr><w:t>x</w:t></w:r></w:p>")
        runs.append(f'<w:p><w:r><w:rPr><w:rStyle w:val="Marked"/>{direct}</w:rPr><w:t>x</w:t></w:r></w:p>')
    package = _styled_package(styles, runs)
    ps = _paragraphs(package)
    for index, direct in enumerate(directs):
        styled, alone = ps[2 * index], ps[2 * index + 1]
        reading = _drawn(package, styled, _runs(styled)[0])
        assert (reading["highlight"], reading["vertAlign"], reading["em"]) == ("yellow", "superscript", "dot"), direct
        assert _drawn(package, alone, _runs(alone)[0])["highlight"] == "none", direct  # a character style's isn't drawn


def test_a_character_style_s_highlight_is_never_drawn() -> None:
    """Round six: alone, with bold, through ``w:basedOn``; a paragraph style's and a direct one are; its shading is."""
    from docx4j_py.wml import read_run_options

    styles = (
        '<w:style w:type="character" w:styleId="Hi"><w:name w:val="Hi"/><w:rPr><w:b/><w:highlight w:val="yellow"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="00FF00"/></w:rPr></w:style>'
        '<w:style w:type="character" w:styleId="HiChild"><w:name w:val="HiChild"/><w:basedOn w:val="Hi"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="ParaHi"><w:name w:val="ParaHi"/><w:rPr><w:highlight w:val="cyan"/></w:rPr></w:style>'
    )
    package = _styled_package(styles, [
        '<w:p><w:r><w:rPr><w:rStyle w:val="Hi"/></w:rPr><w:t>a</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:rStyle w:val="HiChild"/></w:rPr><w:t>b</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="ParaHi"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Hi"/></w:rPr><w:t>c</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:rStyle w:val="Hi"/><w:highlight w:val="green"/></w:rPr><w:t>d</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:rPr><w:rStyle w:val="Hi"/></w:rPr></w:pPr><w:r><w:t>e</w:t></w:r></w:p>',
    ])
    resolver = package.get_property_resolver()
    ps = _paragraphs(package)
    effective = [resolver.get_effective_rpr(_runs(p)[0].r_pr, p.p_pr) for p in ps]
    assert [read_run_options(e)["highlight_color"] for e in effective] == [None, None, "#00FFFF", "#008000", None]
    assert all(is_on(e.b) for e in effective[:4])
    assert [_drawn(package, p, _runs(p)[0])["fill"] for p in ps[:4]] == ["00FF00"] * 4
    assert resolver.get_effective_paragraph_mark_rpr(ps[4].p_pr).highlight is None  # a mark's character style too
    assert resolver.get_effective_rpr("Hi").highlight.val.value == "yellow"  # the style's own resolution keeps it


def test_the_on_off_words_in_w_u_remove_a_style_s_underline() -> None:
    """Round six: exactly ``true``, ``false``, ``on``, ``off``, ``1`` or ``0`` reads ``none``; ``True`` and ``yes`` don't."""
    styles = '<w:style w:type="character" w:styleId="U"><w:name w:val="U"/><w:rPr><w:u w:val="double"/></w:rPr></w:style>'
    values = ("true", "false", "on", "off", "1", "0", "True", "TRUE", "yes", "bogus")
    package = _styled_package(styles, [
        f'<w:p><w:r><w:rPr><w:rStyle w:val="U"/><w:u w:val="{v}"/></w:rPr><w:t>x</w:t></w:r></w:p>' for v in values
    ])
    read = [_drawn(package, p, _runs(p)[0])["u"] for p in _paragraphs(package)]
    assert read == ["none"] * 6 + ["double"] * 4


def test_a_colour_over_a_style_by_attribute() -> None:
    """Round six's table: an invalid text or underline colour is absent; an invalid fill or border colour is automatic."""
    styles = (
        '<w:style w:type="character" w:styleId="All"><w:name w:val="All"/><w:rPr><w:color w:val="FF0000"/>'
        '<w:u w:val="single" w:color="FF0000"/><w:bdr w:val="single" w:sz="4" w:space="0" w:color="FF0000"/>'
        '<w:shd w:val="pct50" w:color="FF0000" w:fill="FF0000"/></w:rPr></w:style>'
    )
    package = _styled_package(styles, [
        ('<w:p><w:r><w:rPr><w:rStyle w:val="All"/><w:color w:val="F00"/><w:u w:color="F00"/>'
        '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="F00"/><w:shd w:val="pct50" w:color="F00" w:fill="F00"/>'
        "</w:rPr><w:t>x</w:t></w:r></w:p>"),
        '<w:p><w:r><w:rPr><w:rStyle w:val="All"/><w:color w:val="auto"/><w:u w:color="auto"/></w:rPr><w:t>y</w:t></w:r></w:p>',
        ('<w:p><w:r><w:rPr><w:rStyle w:val="All"/><w:color w:val="Auto"/><w:u w:color="navy"/>'
        '<w:shd w:val="pct50" w:color="Navy" w:fill=""/></w:rPr><w:t>z</w:t></w:r></w:p>'),
    ])
    ps = _paragraphs(package)
    first, second, third = (_drawn(package, p, _runs(p)[0]) for p in ps)
    assert (first["color"], first["u-color"], first["border"], first["fill"], first["pattern"]) == (
        "FF0000", "FF0000", "000000", "none", "auto",
    )
    assert (second["color"], second["u-color"]) == ("auto", "auto")  # exactly auto overrides
    assert (third["color"], third["u-color"], third["fill"], third["pattern"]) == ("FF0000", "000080", "none", "000080")


# ---------------------------------------------------------------------------
# what a save writes: the enumeration and colour rows (CR-018 section 4.7)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", PROBES)
def test_word_draws_a_normalised_save_as_it_drew_the_source(name: str, tmp_path: Path) -> None:
    """The oracle over this package's normalising save of each probe: every marker drawn as Word drew the probe."""
    from docx4j_py.openpackaging.api import load

    package, _source, _written = _saved(ONOFF / f"{name}.docx")
    path = tmp_path / f"{name}.docx"
    path.write_bytes(package.save())
    ours, word = load(path), Probe(name)
    ps = _paragraphs(ours)
    for m in word.markers:
        r = next(r for r in _runs(ps[m.index]) if m.text in _text(r))
        assert _drawn(ours, ps[m.index], r) == _word_drew(m, _drawn(word.saved, m.saved_p, m.saved_r), "drawn"), m
    if name.startswith("colour-background"):
        assert _page_colour(ours) == _page_colour(word.saved)


def _phase_b_rows(name: str) -> list[tuple[str, str, Any]]:
    package, _source, _written = _saved(ONOFF / f"{name}.docx")
    return [(r.path.rsplit("/", 1)[-1] if "@" in r.path.rsplit("/", 1)[-1] else "element", r.before, r.after)
            for r in package.last_save.normalised if not _on_off_row(r)]


def test_the_enumeration_and_colour_rows() -> None:
    """Each kind of rewrite, as the probes hold it (CR-018 section 4.7's table, rounds two to seven)."""
    rows = {name: _phase_b_rows(name) for name in PROBES}
    assert rows["props-enum-other"] == [
        ("element", '<w:highlight w:val="YELLOW"/>', None),
        ("element", '<w:highlight w:val="bogus"/>', None),
        ("element", '<w:vertAlign w:val="Superscript"/>', None),
        ("element", '<w:vertAlign w:val="bogus"/>', None),
        ("element", '<w:em w:val="DOT"/>', None),
        ("@w:val", "red", "FF0000"),
        ("@w:val", "ff0000", "FF0000"),
    ]
    assert rows["props-enum-missing"] == [
        ("element", "<w:highlight/>", None), ("element", "<w:vertAlign/>", None), ("element", "<w:em/>", None),
        ("element", "<w:color/>", None),
    ]
    assert rows["props-u-valid"] == [("element", "<w:u/>", None)]  # as Word's save drops it
    assert ("@w:val", "true", "none") in rows["props-u-other"] and ("element", '<w:u w:val="bogus"/>', None) in rows["props-u-other"]
    assert sorted({after for _a, _b, after in rows["props-u-other"]} - {None}) == ["none"]
    assert rows["colour-pattern"] == [
        ("@w:color", "red", "FF0000"), ("@w:color", "Navy", "000080"), ("@w:color", "F00", "auto"),
        ("@w:color", "F00", "auto"), ("@w:color", "Navy", "000080"),
    ]
    assert rows["colour-background-red"] == [("@w:color", "red", "FF0000")]
    assert rows["colour-background-navy"] == [("@w:color", "Navy", "000080")]
    assert rows["colour-background-f00"] == [("element", '<w:background w:color="F00"/>', None)]  # as Word drops it
    assert rows["colour-background-hex"] == []
    for name in PROBES[:9]:  # rounds one and two's on/off probes hold no enumeration or colour to rewrite
        assert rows[name] == ([("element", "<w:u/>", None)] if name == "props-b-styles" else []), name
    assert sum(len(found) for found in rows.values()) == 202


def test_a_rejected_colour_beside_a_theme_colour_is_auto() -> None:
    """``w:val`` is required, and the theme colour is what Word draws: a rejected value beside one is written ``auto``."""
    import io
    import zipfile

    package = _styled_package("", [
        '<w:p><w:r><w:rPr><w:color w:val="bogus" w:themeColor="accent1"/></w:rPr><w:t>T1_THEME</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:color w:themeColor="accent2"/></w:rPr><w:t>T2_THEME_ONLY</w:t></w:r></w:p>',
    ])
    xml = zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml").decode()
    assert '<w:color w:val="auto" w:themeColor="accent1"/>' in xml
    assert '<w:color w:themeColor="accent2"/>' in xml  # nothing to read: left alone, as docx4j leaves it
    assert [(r.before, r.after) for r in package.last_save.normalised] == [("bogus", "auto")]


def test_a_revision_s_former_properties_are_normalised_too() -> None:
    """docx4j's ``aRevisionsFormerPropertiesToo``: ``w:rPrChange/w:rPr``, which Reject All restores, run's and mark's."""
    import io
    import zipfile

    rev = ' w:author="A" w:date="2026-10-24T00:00:00Z"'
    package = _styled_package("", [
        (f'<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1"{rev}><w:rPr><w:color w:val="F00"/>'
        '<w:highlight w:val="YELLOW"/><w:vertAlign/><w:em w:val="DOT"/><w:u w:val="single" w:color="red"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="bogus"/></w:rPr></w:rPrChange></w:rPr><w:t>T2_REVISION</w:t></w:r></w:p>'),
        (f'<w:p><w:r><w:rPr><w:i/><w:rPrChange w:id="2"{rev}><w:rPr><w:color w:val="red"/>'
        '<w:highlight w:val="green"/></w:rPr></w:rPrChange></w:rPr><w:t>T3_REVISION_NAME</w:t></w:r></w:p>'),
        (f'<w:p><w:pPr><w:rPr><w:b/><w:rPrChange w:id="3"{rev}><w:rPr><w:color w:val="navy"/>'
        '<w:highlight w:val="bogus"/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:r><w:t>T4_MARK_REVISION</w:t></w:r></w:p>'),
    ])
    xml = zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml").decode()

    def run(marker: str) -> str:
        at = xml.index(marker)
        return xml[xml.rindex("<w:p>", 0, at):at]

    t2 = run("T2_REVISION")
    assert "<w:color" not in t2 and "<w:highlight" not in t2 and "<w:vertAlign" not in t2 and "<w:em" not in t2
    assert 'w:color="FF0000"' in t2 and 'w:fill="auto"' in t2
    t3 = run("T3_REVISION_NAME")
    assert '<w:color w:val="FF0000"/>' in t3 and '<w:highlight w:val="green"/>' in t3
    t4 = run("T4_MARK_REVISION")
    assert '<w:color w:val="000080"/>' in t4 and "<w:highlight" not in t4
    # the tree keeps what was read
    former = _runs(_paragraphs(package)[1])[0].r_pr.r_pr_change.r_pr  # CT_RPrChange's w:rPr holds lists
    assert [c.val for c in former.color] == ["red"]
    assert [h.val for h in _runs(_paragraphs(package)[0])[0].r_pr.r_pr_change.r_pr.highlight] == ["YELLOW"]


def test_the_enumeration_and_colour_rows_keep_the_tree_and_the_option() -> None:
    """The rows are written into the render only; ``save(normalise=False)`` writes the source's spelling."""
    import io
    import zipfile

    from docx4j_py.openpackaging.api import load

    package = load(ONOFF / "props-enum-other.docx")
    n6 = _runs(_paragraphs(package)[10])[1]
    assert n6.r_pr.color.val == "red"
    normalised = zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml")
    assert n6.r_pr.color.val == "red" and package.main_document_part.contents is not None
    kept = zipfile.ZipFile(io.BytesIO(package.save(normalise=False))).read("word/document.xml")
    assert b'<w:color w:val="red"/>' in kept and b'<w:highlight w:val="YELLOW"/>' in kept
    assert b'<w:color w:val="red"/>' not in normalised and b"YELLOW" not in normalised.split(b"N1_HIGHLIGHT_UPPER")[1][:300]
    assert package.last_save.normalise is False and package.last_save.normalised == []
    # part.xml, outside a save, is the tree as it stands
    assert b'<w:color w:val="red"/>' in package.main_document_part.xml


def test_a_value_set_in_code_is_written_in_the_form_too() -> None:
    """Phase B's rows are found in the render, so a colour or an enumeration set in code is normalised as well."""
    import io
    import zipfile

    from docx4j_py.openpackaging.api import create_package
    from docx4j_py.wml import el, p, r

    package = create_package()
    body = package.main_document_part.contents.body
    body.content.append(p(r("x")))
    run = body.content[-1].content[0]
    run.r_pr = el.rPr(color=el.color(val="navy"), highlight=el.highlight(val="Yellow"))
    xml = zipfile.ZipFile(io.BytesIO(package.save())).read("word/document.xml").decode()
    assert '<w:color w:val="000080"/>' in xml and "<w:highlight" not in xml
    assert [(rec.before, rec.after) for rec in package.last_save.normalised] == [
        ("navy", "000080"), ('<w:highlight w:val="Yellow"/>', None),
    ]


def test_the_facade_s_colour_follows_word_not_python_docx_s_slicing() -> None:
    """CR-018 Q9: five or seven digits and ``" FF0000"`` read as Word reads them; a name as its hex."""
    from docx4j_py.docx import Document
    from docx4j_py.docx.enums import MSO_COLOR_TYPE

    document = Document(str(ONOFF / "colour-hex.docx"))
    by_marker = {run.text: run.font.color for p in document.paragraphs for run in p.runs if "_" in run.text}
    read = {text: (None if c.rgb is None else str(c.rgb), c.type) for text, c in by_marker.items()}
    five = next(v for k, v in read.items() if "FF000" in k and "FF0000" not in k and "FF00000" not in k)
    assert five == (None, None)
    seven = next(v for k, v in read.items() if "FF00000" in k)
    assert seven == (None, None)
    assert any(v == ("FF0000", MSO_COLOR_TYPE.RGB) for k, v in read.items() if "#" in k or "HASH" in k.upper())
    names = Document(str(ONOFF / "colour-case.docx"))
    colours = {run.text: run.font.color for p in names.paragraphs for run in p.runs if "_" in run.text}
    assert str(colours["K4_Red"].rgb) == "FF0000" and colours["K4_Red"].type is MSO_COLOR_TYPE.RGB
    assert colours["K1_auto"].rgb is None and colours["K1_auto"].type is MSO_COLOR_TYPE.AUTO


def test_the_fixtures_are_docx4j_s_at_5163c2986() -> None:
    """Every probe and Word's save of it is its git blob in docx4j's ``runprops/word365/`` at ``5163c2986``.

    Rounds one to four were copied from docx4j at ``868a1797b``, five and six from
    the share (and are docx4j's at ``1d1b824a9``), seven from docx4j at
    ``5163c2986`` (its CR-061 phase 3), which holds all seven rounds.
    """
    import subprocess

    docx4j = ROOT.parent / "docx4j"
    if not (docx4j / ".git").exists():
        pytest.skip("../docx4j is not checked out")
    folder = "docx4j-core-tests/src/test/resources/runprops/word365"
    try:
        listing = subprocess.run(
            ["git", "-C", str(docx4j), "ls-tree", "5163c2986", f"{folder}/"],
            check=True, capture_output=True, text=True,
        ).stdout.splitlines()
    except subprocess.CalledProcessError:
        pytest.skip("../docx4j has no commit 5163c2986")
    blobs = {line.rsplit("/", 1)[-1]: line.split()[2] for line in listing if line.endswith(".docx")}
    ours = sorted(p.name for p in ONOFF.glob("*.docx"))
    assert sorted(blobs) == ours and len(ours) == 66
    for name in ours:
        data = (ONOFF / name).read_bytes()
        git_blob = hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest()
        assert git_blob == blobs[name], name


def test_concurrent_saves_keep_their_own_reports() -> None:
    """The render pass finds its save through a context variable: saves in threads do not mix their rows."""
    from concurrent.futures import ThreadPoolExecutor

    from docx4j_py.openpackaging.api import load

    def save(name: str) -> tuple[str, list[tuple[str, Any]]]:
        package = load(ONOFF / f"{name}.docx")
        _ = package.main_document_part.contents
        package.save()
        return name, sorted((r.before, r.after) for r in package.last_save.normalised)

    names = ["props-enum-other", "colour-pattern", "colour-background-red", "props-u-valid", "onoff-valid"]
    alone = dict(save(name) for name in names)
    with ThreadPoolExecutor(8) as pool:
        together = list(pool.map(save, names * 6))
    assert all(rows == alone[name] for name, rows in together)
    assert len(alone["colour-pattern"]) == 5 and len(alone["onoff-valid"]) == 2
