"""Run properties read as Word 365 reads them: CR-018's on/off half (Phase A), over Word's own saves.

The fixtures are ``tests/fixtures/onoff-fun/`` (``SHA1SUMS``; ``tests/README.md``
has their provenance): the twenty probes of rounds one to four, as docx4j's
``runprops/word365/`` holds them at ``868a1797b``, and the eight of rounds five
and six from the share, each with Word 365's save of it (``<name>-word365.docx``).

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
``w:vanish``), and ``format_span``'s write across paragraphs. Phase B adds the
enumerations and colour; where Word's own save changes what Word drew (docx4j
CR-060 section 13 item 5), Word's answer comes from the PDF and is stated in
:data:`WORD_DREW`, none of which is an on/off reading.
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

#: Rounds one to four (docx4j's ``RunPropertiesWord365Test.FILES``), then rounds five and six.
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
    assert len([n for n in listed if n.endswith(".docx")]) == 2 * len(PROBES) == 56


def test_every_probe_has_markers(probes: dict[str, Probe]) -> None:
    counts = {name: len(probe.markers) for name, probe in probes.items()}
    assert counts["onoff-valid"] == 8
    assert counts["props-b-case"] == 10
    assert counts["props-toggles-1"] == counts["props-toggles-2"] == 26  # five values of five properties, and a control
    assert sum(counts[name] for name in PROBES[:20]) == 227  # rounds one to four: CR-018 section 2's 227
    assert sum(counts[name] for name in PROBES[20:]) == 149  # rounds five and six


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
    # the tree keeps what was read
    kept = [k.text for _p, k in package.kept if k.field_name == "val" and k.parent_class == "BooleanDefaultTrue"]
    for record in package.last_save.normalised:
        assert record.before in kept or record.before in {k.text for _p, k in package.kept}


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


@pytest.mark.parametrize("name", sorted(NORMALISED))
def test_the_on_off_rows(name: str) -> None:
    package, _source, _written = _saved(ONOFF / f"{name}.docx")
    found = {record.before: record.after for record in package.last_save.normalised}
    assert found == NORMALISED[name]
    assert all(record.part == "/word/document.xml" for record in package.last_save.normalised)
    assert all(record.path.endswith("/w:rPr/w:b/@w:val") for record in package.last_save.normalised)


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
