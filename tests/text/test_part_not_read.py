"""``part.not_read``: a styles, numbering or settings part the reading needs and cannot read. CR-031 sections 4.6, 5.

WD 0.66 REQ-801 (the user, 2026-11-03: "part.not_read as recommended"); docx4j
CR-046 section 10 at ``3ad164bdf``, its ``Extractor.consult``. A part the
reading consults and does not emit whose XML is not well formed is read as if
the package lacked it, and the result carries one ``part.not_read``, its
selector the part's name, raised where the reading first needs the part: the
styles at the first paragraph or run resolved, the numbering at the first
paragraph with a ``w:numId``, the settings for ``fields_update_on_open`` (every
reading but one level down), forms protection and, where a section has an even
header or footer, ``w:evenAndOddHeaders``. The font table is never consulted.

**docx4j's fixture** (``tests/fixtures/part-not-read/``, copied at ``3ad164bdf``):
each file in each view held to ``expected/<fixture>.<view>.txt`` and
``expected-warnings.tsv``, and its ``PartNotReadTest``. Its expected text maps
case (``w:caps`` by the run's language, ``w:smallCaps`` under ``small_caps:
upper``), which waits for CR-029 here: those lines are pending, each held to be
the case mapping of what this package reads, and every other line exact.

**Section 4.6's measurement**, as a test: a document made in code, each of its
styles, numbering and settings parts cut to two thirds and removed, the cut
reading as the removed with one ``part.not_read``; the font table not consulted.
"""

from __future__ import annotations

import hashlib
import io
import logging
import re
import zipfile
from pathlib import Path

import pytest

from docx4j_py import create_package, load
from docx4j_py.fragments import wml
from docx4j_py.model.listnumbering.emulator import Emulator
from docx4j_py.model.text import TextResult, TextWarning, extract, markdown, plain_text
from docx4j_py.model.text.extractor import EMPTY_FORM_FIELD
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.parts.binary_part import EmbeddedPackagePart
from docx4j_py.openpackaging.parts.wml import FontTablePart, HeaderPart, NumberingDefinitionsPart

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "part-not-read"
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PARTS = ("styles", "numbering", "settings")


def _unzip(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {name: z.read(name) for name in z.namelist()}


def _zip(entries: dict[str, bytes]) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)
    return out.getvalue()


def _saved(pkg) -> bytes:
    out = io.BytesIO()
    pkg.save(out)
    return out.getvalue()


def _warnings(result: TextResult) -> str:
    """docx4j's ``warnings``: kind and selector, ``; `` between two."""
    return "; ".join(f"{w.kind} {w.selector}" for w in result.warnings)


# ---------------------------------------------------------------------------
# docx4j's fixture: the copy, and every file in every view
# ---------------------------------------------------------------------------


def test_the_copy_is_docx4js():
    """34 files, each to ``SHA1SUMS`` (docx4j's git blobs at ``3ad164bdf``); our README is not among them."""
    listed = {}
    for line in (FIXTURES / "SHA1SUMS").read_text().splitlines():
        digest, name = line.split(maxsplit=1)
        listed[name] = digest
    on_disk = {p.relative_to(FIXTURES).as_posix() for p in FIXTURES.rglob("*") if p.is_file()}
    assert set(listed) == on_disk - {"SHA1SUMS", "README.md"}
    assert len(listed) == 34
    for name, digest in listed.items():
        assert hashlib.sha1((FIXTURES / name).read_bytes()).hexdigest() == digest, name


ROWS = [row.split("\t") for row in (FIXTURES / "expected-warnings.tsv").read_text().splitlines()]
assert ROWS[0] == ["fixture", "view", "warnings"]

#: The lines of docx4j's expected text that map case, which CR-029 builds here (0-based): the run in capitals by its
#: character style, the paragraph in small capitals by its style (under ``small_caps: upper``), the run in direct
#: ``w:caps``. Where docx4j's text differs from this package's on one of them, it is pending.
CASE_LINES = {5: "istanbul caps-by-style", 6: "small caps by style", 7: "istanbul direct caps"}

#: The pending lines of each fixture, in each view: docx4j maps case where the styles say so, by the settings'
#: ``themeFontLang`` tr-TR ("İSTANBUL") where the settings can be read, else the default language ("ISTANBUL").
PENDING = {
    "part-not-read-intact.docx": 3,
    "part-not-read-numbering.docx": 3,
    "part-not-read-numbering-absent.docx": 3,
    "part-not-read-settings.docx": 3,
    "part-not-read-settings-absent.docx": 3,
    "part-not-read-styles.docx": 1,  # without styles only the direct w:caps maps
    "part-not-read-styles-absent.docx": 1,
    "part-not-read-numbering-unneeded.docx": 0,
}


def _mapped(text: str, turkish: bool) -> str:
    """CR-029's capitals for these lines: Unicode's upper case, with the Turkish dotted capital I under tr-TR."""
    return (text.replace("i", "İ") if turkish else text).upper()


def _read(fixture: str, view: str) -> TextResult:
    # docx4j's options: list_labels on, small_caps upper (CR-029's option, not here yet), the rest the defaults
    return extract(load(FIXTURES / fixture), list_labels=True, view=view)


@pytest.mark.parametrize(("fixture", "view", "expected_warnings"), ROWS[1:], ids=[f"{r[0]}-{r[1]}" for r in ROWS[1:]])
def test_each_fixture_reads_as_docx4j_expects(fixture: str, view: str, expected_warnings: str):
    """docx4j's ``eachFixtureReadsAsExpected``: the warnings exactly, the text exactly but for CR-029's case lines."""
    result = _read(fixture, view)
    assert _warnings(result) == expected_warnings
    got = plain_text(result).splitlines()
    expected = (FIXTURES / "expected" / f"{fixture.removesuffix('.docx')}.{view}.txt").read_text(encoding="utf-8")
    want = expected.splitlines()
    assert len(got) == len(want)
    turkish = "settings" not in fixture  # the settings' themeFontLang tr-TR, where they can be read
    pending = []
    for i, (line, line_wanted) in enumerate(zip(got, want, strict=True)):
        if line == line_wanted:
            continue
        # pending until CR-029: the case-mapping run, read as stored here, its expected text its capitals
        assert i in CASE_LINES and line == CASE_LINES[i], (i, line, line_wanted)
        assert line_wanted == _mapped(line, turkish), (i, line, line_wanted)
        pending.append(i)
    assert len(pending) == PENDING[fixture]  # when CR-029 maps case, this fails: the pending lines go


def test_the_fixture_has_twenty_four_rows():
    assert len(ROWS) - 1 == 8 * 3


@pytest.mark.parametrize("view", ["accepted", "original", "markup"])
@pytest.mark.parametrize("part", PARTS)
def test_a_part_that_cannot_be_read_reads_as_a_part_that_is_absent(part: str, view: str):
    """docx4j's ``aPartThatCannotBeReadReadsAsAPartThatIsAbsent``: the same text, and one ``part.not_read``."""
    broken = _read(f"part-not-read-{part}.docx", view)
    absent = _read(f"part-not-read-{part}-absent.docx", view)
    assert plain_text(broken) == plain_text(absent)
    assert _warnings(broken) == f"part.not_read /word/{part}.xml"
    assert _warnings(absent) == ""
    (warning,) = broken.warnings
    assert warning.kind == TextWarning.PART_NOT_READ == "part.not_read"
    assert warning.address is None and warning.part_name == f"/word/{part}.xml"
    assert warning.to_dict() == {"kind": "part.not_read", "address": None, "selector": f"/word/{part}.xml"}
    assert str(warning) == f"part.not_read (/word/{part}.xml)"


def test_the_emulator_counts_direct_numbering_without_readable_styles():
    """docx4j's ``theEmulatorCountsDirectNumberingWithoutReadableStyles`` (its fix (a), which holds here already):
    outside the text model, a paragraph's own numbering is counted where the styles part cannot be read."""
    pkg = load(FIXTURES / "part-not-read-styles.docx")
    body = pkg.main_document_part.contents.body.content
    first = Emulator.get_number(pkg, body[2].p_pr)
    assert first is not None and first.num_string == "1."
    assert Emulator.get_number(pkg, body[3].p_pr).num_string == "2."
    assert Emulator.get_number(pkg, body[0].p_pr) is None  # numbered by its style, which is lost


def test_the_unreadable_part_is_kept():
    """docx4j's ``theUnreadablePartIsKept``: the reading changes nothing; an untouched save keeps the part's bytes."""
    fixture = FIXTURES / "part-not-read-styles.docx"
    pkg = load(fixture)
    extract(pkg, list_labels=True)
    assert _unzip(_saved(pkg))["word/styles.xml"] == _unzip(fixture.read_bytes())["word/styles.xml"]


def _embedding(inner: bytes, name: str):
    outer = create_package()
    mdp = outer.main_document_part
    embedded = EmbeddedPackagePart(f"/word/embeddings/{name}", DOCX)
    embedded.set_bytes(inner)
    rel_id = mdp.add_target_part(embedded).id
    mdp.contents.body.content.append(
        wml(
            f'<w:p {W} xmlns:o="urn:schemas-microsoft-com:office:office" {R}><w:r><w:object>'
            f'<o:OLEObject Type="Embed" ProgID="Word.Document.12" r:id="{rel_id}"/></w:object></w:r></w:p>'
        )
    )
    return outer


def test_inside_an_embedded_document_the_warning_names_its_part():
    """docx4j's ``insideAnEmbeddedDocumentTheWarningNamesItsPart``: one level down, the embedded package's part."""
    outer = _embedding((FIXTURES / "part-not-read-styles.docx").read_bytes(), "styles-not-read.docx")
    result = extract(outer, list_labels=True, embedded=True)
    assert _warnings(result) == "part.not_read /word/embeddings/styles-not-read.docx"
    assert "1.\tfirst item" in plain_text(result)


def test_one_level_down_the_settings_are_not_needed_for_fields_update_on_open():
    """docx4j's ``!nested && fieldsUpdateOnOpen()``: a nested reading's updateFields is not the result's, so an embedded
    document's unreadable settings, which nothing else there needs, raise nothing."""
    outer = _embedding((FIXTURES / "part-not-read-settings.docx").read_bytes(), "settings-not-read.docx")
    result = extract(outer, list_labels=True, embedded=True)
    assert _warnings(result) == ""
    assert "1.\tfirst item" in plain_text(result)
    assert result.metadata[TextResult.FIELDS_UPDATE_ON_OPEN] is False


def test_a_part_s_reading_consults_the_settings_too():
    """``extract(part)``, docx4j's ``Texts.of(Part)``: its result carries ``fields_update_on_open`` as well."""
    pkg = load(FIXTURES / "part-not-read-settings.docx")
    result = extract(pkg.main_document_part, list_labels=True)
    assert _warnings(result) == "part.not_read /word/settings.xml"
    assert result.metadata[TextResult.FIELDS_UPDATE_ON_OPEN] is False


@pytest.mark.parametrize("list_labels", [False, True])
def test_a_numbering_part_no_paragraph_needs_raises_nothing(list_labels: bool):
    """REQ-801: a part the document does not need raises nothing, whatever the options (``numbering-unneeded``)."""
    pkg = load(FIXTURES / "part-not-read-numbering-unneeded.docx")
    for view in ("accepted", "original", "markup"):
        result = extract(pkg, list_labels=list_labels, view=view)
        assert result.warnings == []
        assert plain_text(result) == "plain one\nplain two\n"
    assert "plain one" in markdown(extract(pkg))
    assert not pkg.main_document_part.numbering_definitions_part.is_content_unreadable  # never even read


def test_nothing_is_logged_by_the_text_model(caplog: pytest.LogCaptureFixture):
    """Section 4.6: no log line; the warning is the text model's one record of the failure."""
    for part in PARTS:
        caplog.clear()
        with caplog.at_level(logging.DEBUG):
            result = _read(f"part-not-read-{part}.docx", "accepted")
        assert _warnings(result) == f"part.not_read /word/{part}.xml"
        assert [r for r in caplog.records if r.name.startswith("docx4j_py.model.text")] == []
        named = [r for r in caplog.records if r.name.startswith("docx4j_py") and f"/word/{part}.xml" in r.getMessage()]
        assert named == []


def test_the_content_api_refuses_an_unreadable_styles_part():
    """Section 4.6 (D6): the content API writes, so it does not resolve without styles: the resolver and every reader
    that resolves styles raise the styles part's own exception (docx4j: a NullPointerException)."""
    pkg = load(FIXTURES / "part-not-read-styles.docx")
    styles = pkg.main_document_part.style_definitions_part
    with pytest.raises(Docx4JException) as raised:
        pkg.get_property_resolver()
    assert raised.value is styles.unreadable
    assert str(raised.value).startswith("Problem with part /word/styles.xml: ")
    with pytest.raises(Docx4JException) as again:
        _ = pkg.body.text
    assert again.value is styles.unreadable
    # the text model, a reader, goes on without them
    assert "hidden-by-style" in plain_text(extract(pkg))


# ---------------------------------------------------------------------------
# section 4.6's measurement: each part cut to two thirds reads as the part removed
# ---------------------------------------------------------------------------

_P = "<w:p {w}>{ppr}<w:r>{rpr}<w:t xml:space='preserve'>{text}</w:t></w:r></w:p>"


def _p(text: str, ppr: str = "", rpr: str = "") -> object:
    return wml(
        _P.format(
            w=W, ppr=f"<w:pPr>{ppr}</w:pPr>" if ppr else "", rpr=f"<w:rPr>{rpr}</w:rPr>" if rpr else "", text=text
        )
    )


FORM = (
    f"<w:p {W}><w:r><w:fldChar w:fldCharType='begin'><w:ffData><w:name w:val='Text1'/><w:enabled/>"
    "<w:calcOnExit w:val='0'/><w:textInput/></w:ffData></w:fldChar></w:r>"
    "<w:r><w:instrText xml:space='preserve'> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType='separate'/></w:r>"
    "<w:r><w:t>TYPED</w:t></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r></w:p>"
)


def _measured(numbered: int = 1) -> dict[str, bytes]:
    """Section 4.6's document: a heading numbered by its style, `numbered` paragraphs numbered directly, a paragraph
    hidden by its style, a run in Symbol, a plain paragraph and a FORMTEXT; a default and an even header under
    ``w:evenAndOddHeaders``, ``w:updateFields`` on and forms protection enforced; a font table."""
    pkg = create_package()
    mdp = pkg.main_document_part
    body = mdp.contents.body
    body.content.append(_p("HEADING BY STYLE", "<w:pStyle w:val='Heading1'/>"))
    for n in range(numbered):
        body.content.append(_p(f"DIRECT LIST ITEM {n}", "<w:numPr><w:ilvl w:val='0'/><w:numId w:val='1'/></w:numPr>"))
    body.content.append(_p("HIDDEN BY STYLE", "<w:pStyle w:val='HiddenStyle'/>"))
    body.content.append(_p("abg SYMBOL RUN", rpr="<w:rFonts w:ascii='Symbol' w:hAnsi='Symbol'/>"))
    body.content.append(_p("PLAIN PARAGRAPH"))
    body.content.append(wml(FORM))
    styles = mdp.style_definitions_part.contents
    styles.style.append(
        wml(
            f"<w:style {W} w:type='paragraph' w:styleId='HiddenStyle'><w:name w:val='Hidden Style'/><w:rPr><w:vanish/></w:rPr></w:style>"
        )
    )
    styles.style.append(
        wml(
            f"<w:style {W} w:type='paragraph' w:styleId='Heading1'><w:name w:val='heading 1'/>"
            "<w:pPr><w:numPr><w:ilvl w:val='0'/><w:numId w:val='2'/></w:numPr></w:pPr></w:style>"
        )
    )
    numbering = NumberingDefinitionsPart()
    mdp.add_target_part(numbering)
    numbering.unmarshal_default_numbering()
    first = numbering.contents.num[0].abstract_num_id.val
    numbering.contents.num.append(wml(f"<w:num {W} w:numId='2'><w:abstractNumId w:val='{first}'/></w:num>"))
    for n, (kind, text) in enumerate((("default", "ODD HEADER"), ("even", "EVEN HEADER")), 1):
        header = HeaderPart(f"/word/header{n}.xml")
        rel = mdp.add_target_part(header)
        header.set_contents(wml(f"<w:hdr {W}><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:hdr>"))
        body.sect_pr.header_reference_or_footer_reference.append(
            wml(f"<w:headerReference {W} {R} w:type='{kind}' r:id='{rel.id}'/>")
        )
    settings = mdp.document_settings_part.contents
    settings.even_and_odd_headers = wml(f"<w:evenAndOddHeaders {W}/>")
    settings.update_fields = wml(f"<w:updateFields {W} w:val='true'/>")
    settings.document_protection = wml(f"<w:documentProtection {W} w:edit='forms' w:enforcement='1'/>")
    fonts = FontTablePart()
    mdp.add_target_part(fonts)
    fonts.unmarshal_default_fonts()
    return _unzip(_saved(pkg))


NAMES = {
    "styles": "word/styles.xml",
    "numbering": "word/numbering.xml",
    "settings": "word/settings.xml",
    "fontTable": "word/fontTable.xml",
}


def _variant(entries: dict[str, bytes], which: str, how: str) -> bytes:
    """`entries` with `which` cut to two thirds of its bytes, or removed (part, relationship and override)."""
    entries = dict(entries)
    name = NAMES[which]
    if how == "cut":
        entries[name] = entries[name][: len(entries[name]) * 2 // 3]
        return _zip(entries)
    del entries[name]
    target = name.split("/", 1)[1]
    rels = entries["word/_rels/document.xml.rels"].decode()
    rels, count = re.subn(rf'<Relationship [^>]*Target="{re.escape(target)}"[^>]*/>', "", rels)
    assert count == 1
    entries["word/_rels/document.xml.rels"] = rels.encode()
    types = entries["[Content_Types].xml"].decode()
    types, count = re.subn(rf'<Override [^>]*PartName="/{re.escape(name)}"[^>]*/>', "", types)
    assert count == 1
    entries["[Content_Types].xml"] = types.encode()
    return _zip(entries)


def _reading(data: bytes) -> TextResult:
    return extract(load(data), list_labels=True, include="headers", fields="update")


def _lines(result: TextResult) -> list[str]:
    return [((b.list_label + " ") if b.list_label else "") + b.text for b in result.blocks if b.text or b.list_label]


MEASURED = _measured()


def test_the_measured_document_reads_whole():
    result = _reading(_zip(MEASURED))
    lines = _lines(result)
    assert result.warnings == []
    assert "ODD HEADER" in lines and "EVEN HEADER" in lines
    assert "HIDDEN BY STYLE" not in lines
    assert sum(1 for line in lines if line.startswith("1.")) == 1 and any(line.startswith("2.") for line in lines)
    assert "TYPED" in lines  # forms protection on: what was typed is kept
    assert result.metadata[TextResult.FIELDS_UPDATE_ON_OPEN] is True


@pytest.mark.parametrize("which", PARTS)
def test_each_cut_part_reads_as_the_part_removed(which: str):
    """Section 4.6's table, as a test: the same reading, with one ``part.not_read`` naming the part."""
    cut = _reading(_variant(MEASURED, which, "cut"))
    removed = _reading(_variant(MEASURED, which, "absent"))
    assert _lines(cut) == _lines(removed)
    assert cut.metadata == removed.metadata
    assert [(w.kind, w.address, w.part_name) for w in cut.warnings] == [
        (TextWarning.PART_NOT_READ, None, f"/{NAMES[which]}")
    ]
    assert removed.warnings == []
    lines = _lines(cut)
    if which == "styles":
        # the heading's label gone and the hidden text shown; the direct label kept (1., the heading no longer counted)
        assert "HEADING BY STYLE" in lines and "HIDDEN BY STYLE" in lines
        assert "1.\t DIRECT LIST ITEM 0" in lines
    elif which == "numbering":
        assert not [b for b in cut.blocks if b.list_label]  # every label gone, by style and direct
    else:
        assert "EVEN HEADER" not in lines  # w:evenAndOddHeaders lost
        assert cut.metadata[TextResult.FIELDS_UPDATE_ON_OPEN] is False
        assert EMPTY_FORM_FIELD in lines and "TYPED" not in lines  # forms protection off: Word's update applies


def test_the_font_table_is_never_consulted():
    """Section 4.6: nothing under ``model/text/`` reads one, so a font table not well formed raises nothing."""
    intact = _reading(_zip(MEASURED))
    cut = _reading(_variant(MEASURED, "fontTable", "cut"))
    assert _lines(cut) == _lines(intact) and cut.warnings == []


def test_one_warning_a_part_however_often_the_reading_needs_it():
    """Once a part: twenty numbered paragraphs, both headers, every view and markdown, all three parts broken; the
    settings first (``fields_update_on_open``), then the styles (the first paragraph), then the numbering."""
    entries = _measured(numbered=20)
    for which in PARTS:
        entries[NAMES[which]] = entries[NAMES[which]][: len(entries[NAMES[which]]) * 2 // 3]
    pkg = load(_zip(entries))
    for view in ("accepted", "original", "markup"):
        result = extract(pkg, list_labels=True, include="headers", view=view)
        assert [w.part_name for w in result.warnings] == [
            "/word/settings.xml",
            "/word/styles.xml",
            "/word/numbering.xml",
        ]
    result = extract(pkg, list_labels=True, include="headers")
    assert markdown(result).count("1.") <= 1  # no list labels: the numbering was not read
    for which in PARTS:
        assert pkg.main_document_part.package.get_part("/" + NAMES[which]).is_content_unreadable


def test_a_read_refused_for_another_reason_is_not_part_not_read():
    """Section 4.6: only a part whose content could not be read (the engine's record, ``unreadable``) raises it; a
    strict load refusing a styles part for its skipped content reads without styles, silently, as before."""
    from docx4j_py.openpackaging.load import LoadOptions

    entries = _unzip((FIXTURES / "part-not-read-intact.docx").read_bytes())
    entries["word/styles.xml"] = entries["word/styles.xml"].replace(b"</w:styles>", b"<w:bogus/></w:styles>")
    pkg = load(_zip(entries), options=LoadOptions(strict=True))
    result = extract(pkg, list_labels=True)
    assert result.warnings == []
    styles = pkg.main_document_part.style_definitions_part
    assert not styles.is_content_unreadable
    assert "visible hidden-by-style end" in plain_text(result)  # no styles: the character style's w:vanish not read


def test_a_well_formed_numbering_part_is_consulted_without_a_typed_read():
    """The cost: nothing in the text model reads the numbering typed (the emulator reads it with lxml), so the consult
    checks it with a parse that builds no tree, and reads it typed only where that fails, for the engine's record.
    Measured on ``NumberingIndents.docx``: a first reading 18.5 ms before CR-031, 25.8 ms with a typed read, 19.0 ms
    with the check (best of fifteen)."""
    pkg = load(FIXTURES / "part-not-read-intact.docx")
    numbering = pkg.main_document_part.numbering_definitions_part
    result = extract(pkg, list_labels=True)
    assert "1.\tfirst item" in plain_text(result)
    assert not numbering.is_unmarshalled and numbering._private_contents is None
    assert numbering.unreadable is None
    broken = load(FIXTURES / "part-not-read-numbering.docx")
    extract(broken, list_labels=True)
    assert broken.main_document_part.numbering_definitions_part.is_content_unreadable  # read, for the record
