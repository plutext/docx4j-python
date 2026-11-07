"""A part whose XML is not well formed is refused when read, not silently mended. CR-031 sections 4 and 5 (Phase A).

The fork's lxml handler recovered (``etree.iterparse(recover=True)``, upstream
xsdata's setting): a cut part read as far as the cut, a stray ``&`` or an
undefined entity dropped, a second root lost, a cut relationships part loaded
without the parts the rest named, all with an empty skipped report, even under
``LoadOptions(strict=True)`` (CR-031 section 2.2). Since the fork's stage 15
(docx4j-xsdata 26.2.5) the engine parses with ``recover=False``
(``runtime.RECOVER``), as docx4j does (its CR-065 D1).

**The cases** are section 2.2's twenty, each a replacement in the saved bytes of
one document this package makes (docx4j's ``fixture``): four paragraphs, the
second numbered from docx4j's default numbering and the third hidden by its
paragraph style, and a header. Over every case and at both ``strict`` settings:

- ``load()`` returns, except for a relationships part (cases 17 to 19), which
  refuses the load naming it (D5), and the content types (case 20), refused as
  before;
- ``contents`` and ``read_contents()`` raise the part's ``Docx4JException``,
  "Problem with part <name>: <lxml's message>", lxml's ``XMLSyntaxError`` the
  deepest cause (D2, Q8); ``source_tree()`` gives None (D9, section 4.9);
- ``unreadable`` and ``is_content_unreadable`` record it, and a second read
  raises it again without a parse, until the settings change (D8);
- an untouched save, a save after a read of another part and a save after an
  edit of another part each write the broken part byte for byte (D3);
- ``set_contents`` replaces it, and the save writes the replacement;
- ``misplaced()`` reports ``part.not_well_formed`` and goes on (D9);
- ``mce_preprocess=True`` and ``mc_load_mode="word"`` raise the same, with and
  without a block-level ``mc:AlternateContent`` (D2);
- a custom XML part's ``tree`` raises naming its part (D3, Q9).

Then docx4j's ``UnreadablePartTest``'s five engine tests over its own breaks
(``</w:body``, ``</w:hdr``, a ``w:defaultTabStop`` left open), the sniffs over
the corpus (D4), and the census: no parser in the package recovers (D4).
"""

from __future__ import annotations

import ast
import io
import re
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any, NamedTuple

import pytest
from lxml import etree

from docx4j_py import create_package, load
from docx4j_py.fragments import wml
from docx4j_py.openpackaging.exceptions import Docx4JException, InvalidFormatException
from docx4j_py.openpackaging.load import LoadOptions
from docx4j_py.openpackaging.parts.default_xml_part import CustomXmlDataStoragePart
from docx4j_py.openpackaging.parts.docprops import CustomXmlDataStoragePropertiesPart
from docx4j_py.openpackaging.parts.wml import HeaderPart, NumberingDefinitionsPart
from docx4j_py.openpackaging.parts.xml_part import XmlPart, _root_ns_map, syntax_error_of
from docx4j_py.validator import PartNotWellFormed

ROOT = Path(__file__).resolve().parents[2]
SAMPLES = ROOT / "samples"
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
DOC = "word/document.xml"


# ---------------------------------------------------------------------------
# the document, and section 2.2's twenty cases
# ---------------------------------------------------------------------------


def _unzip(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {name: z.read(name) for name in z.namelist()}


def _zip(entries: dict[str, bytes]) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)
    return out.getvalue()


def _saved(pkg: Any) -> bytes:
    out = io.BytesIO()
    pkg.save(out)
    return out.getvalue()


def _base() -> dict[str, bytes]:
    """CR-031 section 2.2's document, made in code and saved: four paragraphs and a header."""
    pkg = create_package()
    mdp = pkg.main_document_part
    body = mdp.contents.body
    for xml in (
        f'<w:p {W}><w:pPr><w:spacing w:before="120"/></w:pPr><w:r><w:t>FIRST PARAGRAPH</w:t></w:r></w:p>',
        (
            f'<w:p {W}><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
            "<w:r><w:t>SECOND PARAGRAPH</w:t></w:r></w:p>"
        ),
        f'<w:p {W}><w:pPr><w:pStyle w:val="HiddenStyle"/></w:pPr><w:r><w:t>HIDDEN BY STYLE</w:t></w:r></w:p>',
        f"<w:p {W}><w:r><w:t>THIRD PARAGRAPH</w:t></w:r></w:p>",
    ):
        body.content.append(wml(xml))
    mdp.style_definitions_part.contents.style.append(
        wml(
            f'<w:style {W} w:type="paragraph" w:styleId="HiddenStyle"><w:name w:val="Hidden Style"/>'
            "<w:rPr><w:vanish/></w:rPr></w:style>"
        )
    )
    numbering = NumberingDefinitionsPart()
    mdp.add_target_part(numbering)
    numbering.unmarshal_default_numbering()
    header = HeaderPart()
    rel = mdp.add_target_part(header)
    header.set_contents(wml(f"<w:hdr {W}><w:p><w:r><w:t>HEADER TEXT</w:t></w:r></w:p></w:hdr>"))
    body.sect_pr.header_reference_or_footer_reference.append(
        wml(f'<w:headerReference {W} {R} w:type="default" r:id="{rel.id}"/>')
    )
    return _unzip(_saved(pkg))


BASE = _base()


def _cut_after(marker: bytes, keep: int) -> Callable[[bytes], bytes]:
    def cut(data: bytes) -> bytes:
        assert marker in data, marker
        return data[: data.index(marker) + keep]

    return cut


def _replace(old: bytes, new: bytes) -> Callable[[bytes], bytes]:
    def replace(data: bytes) -> bytes:
        assert old in data, old
        return data.replace(old, new, 1)

    return replace


def _append(extra: bytes) -> Callable[[bytes], bytes]:
    return lambda data: data + extra


class Case(NamedTuple):
    """One of section 2.2's malformations: the part, the change, and lxml's message as section 2.3 gives it."""

    part: str
    change: Callable[[bytes], bytes]
    #: The start of lxml's message (libxml2's), each line and column after it.
    message: str


CASES: dict[int, Case] = {
    1: Case(DOC, _cut_after(b"SECOND PARAGRAPH", 3), "Premature end of data in tag t line 2"),
    2: Case(DOC, _cut_after(b"<w:pPr><w:pStyle", 9), "Couldn't find end of Start Tag w"),
    3: Case(DOC, _replace(b"SECOND PARAGRAPH", b"SECOND&nbsp;PARAGRAPH"), "Entity 'nbsp' not defined"),
    4: Case(
        DOC,
        _replace(b'<w:spacing w:before="120"/>', b'<w:spacing w:before="120">'),
        "Opening and ending tag mismatch: spacing",
    ),
    5: Case(DOC, _replace(b"</w:body>", b""), "Opening and ending tag mismatch: body"),
    6: Case(DOC, _replace(b"</w:body>", b"</w:body"), "expected '>'"),
    7: Case(DOC, _replace(b"SECOND PARAGRAPH", b"SECOND & PARAGRAPH"), "xmlParseEntityRef: no name"),
    8: Case(DOC, _replace(b"SECOND PARAGRAPH", b"SECOND < PARAGRAPH"), "StartTag: invalid element name"),
    9: Case(
        DOC,
        _replace(b'<w:spacing w:before="120"/>', b'<w:spacing w:before="120" w:before="240"/>'),
        "Attribute w:before redefined",
    ),
    10: Case(DOC, _replace(b"SECOND PARAGRAPH", b"SECOND\x01PARAGRAPH"), "PCDATA invalid Char value 1"),
    11: Case(DOC, _replace(b"SECOND PARAGRAPH", b"SECOND \xe9 PARAGRAPH"), "Invalid bytes in character encoding"),
    12: Case(
        DOC,
        _append(f"<w:document {W}><w:body><w:p><w:r><w:t>EXTRA</w:t></w:r></w:p></w:body></w:document>".encode()),
        "Extra content at the end of the document",
    ),
    13: Case(DOC, _replace(b"<w:document", b"<<w:document"), "StartTag: invalid element name"),
    14: Case("word/header1.xml", _cut_after(b"HEADER TEXT", 3), "Premature end of data in tag t line 2"),
    15: Case("word/styles.xml", _cut_after(b'w:styleId="HiddenStyle"', 0), "Couldn't find end of Start Tag style"),
    16: Case("word/numbering.xml", _cut_after(b"<w:num ", 0), "Premature end of data in tag numbering"),
}

#: Cases 17 to 20: the relationships and the content types, which the load reads.
LOAD_CASES: dict[int, tuple[Case, type[Exception], str]] = {
    17: (
        Case(
            "word/_rels/document.xml.rels", _cut_after(b'Id="rId3"/>', 11), "Premature end of data in tag Relationships"
        ),
        Docx4JException,
        "Problem with part /word/_rels/document.xml.rels: ",
    ),
    18: (
        Case("_rels/.rels", _cut_after(b"<Relationship ", 30), "AttValue: ' expected"),
        Docx4JException,
        "Problem with part /_rels/.rels: ",
    ),
    19: (
        Case("_rels/.rels", _cut_after(b'Id="rId1"/>', 31), "Specification mandates value for attribute Target"),
        Docx4JException,
        "Problem with part /_rels/.rels: ",
    ),
    20: (
        Case("[Content_Types].xml", _cut_after(b"<Override ", 20), "AttValue: ' expected"),
        InvalidFormatException,
        "Bad [Content_Types].xml: ",
    ),
}


def broken(number: int) -> tuple[bytes, dict[str, bytes]]:
    """The package of case `number`, and its entries."""
    case = CASES[number] if number in CASES else LOAD_CASES[number][0]
    entries = dict(BASE)
    entries[case.part] = case.change(BASE[case.part])
    return _zip(entries), entries


def _part(pkg: Any, entry: str) -> XmlPart:
    found = pkg.get_part("/" + entry)
    assert isinstance(found, XmlPart), entry
    return found


def _assert_refused(error: BaseException, part_name: str, message: str) -> etree.XMLSyntaxError:
    """The engine's refusal: ``Docx4JException`` naming the part, lxml's message, lxml's error the deepest cause."""
    assert isinstance(error, Docx4JException)
    syntax = syntax_error_of(error)
    assert isinstance(syntax, etree.XMLSyntaxError)
    deepest = error
    while deepest.__cause__ is not None:
        deepest = deepest.__cause__
    assert deepest is syntax
    # lxml's msg, without the "(<string>, line 2)" its str() appends
    assert str(error) == f"Problem with part {part_name}: {syntax.msg}"
    assert str(syntax.msg).startswith(message), syntax.msg
    assert "(<string>" not in str(error)
    return syntax


def test_the_base_document_reads_whole():
    """The document every case breaks: well formed, read in full, nothing skipped."""
    from docx4j_py.traversal import text_of

    pkg = load(_zip(BASE), options=LoadOptions(strict=True))
    texts = [text_of(p) for p in pkg.main_document_part.contents.body.content]
    assert texts == ["FIRST PARAGRAPH", "SECOND PARAGRAPH", "HIDDEN BY STYLE", "THIRD PARAGRAPH"]
    assert [p.text for p in pkg.body.paragraphs][2] == ""  # hidden by its style
    assert not pkg.skipped


# ---------------------------------------------------------------------------
# every case, at both strict settings
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("strict", [False, True], ids=["lenient", "strict"])
@pytest.mark.parametrize("number", sorted(CASES))
def test_the_load_returns_and_every_read_is_refused(number: int, strict: bool):
    """D2, D8, D9, D10: the load lazy; ``contents``, ``read_contents()`` refused, the failure kept; ``source_tree()`` None."""
    case = CASES[number]
    data, _entries = broken(number)
    pkg = load(data, options=LoadOptions(strict=strict))
    part = _part(pkg, case.part)
    assert part.unreadable is None and not part.is_content_unreadable  # nothing read yet
    with pytest.raises(Docx4JException) as first:
        part.read_contents()
    syntax = _assert_refused(first.value, f"/{case.part}", case.message)
    assert part.unreadable is first.value and part.is_content_unreadable
    with pytest.raises(Docx4JException) as second:
        _ = part.contents
    assert second.value is first.value  # raised again, the same failure
    assert not part.is_unmarshalled and part._private_contents is None
    assert part.skipped == []  # nothing read, nothing reported as skipped
    assert part.source_tree() is None  # a source that held nothing readable
    line, column = syntax.position
    assert f"line {line}, column {column}" in str(syntax.msg)


@pytest.mark.parametrize("number", sorted(CASES))
def test_a_second_read_raises_again_without_a_parse(number: int, monkeypatch: pytest.MonkeyPatch):
    """D8, docx4j's ``unreadable``: a later read raises the kept failure; the part is not parsed again."""
    case = CASES[number]
    pkg = load(broken(number)[0])
    part = _part(pkg, case.part)
    with pytest.raises(Docx4JException) as first:
        _ = part.contents

    def no_parse(*_args: Any, **_kwargs: Any) -> Any:
        raise AssertionError("parsed again")

    monkeypatch.setattr(XmlPart, "_parse", no_parse)
    for read in (lambda: part.contents, part.read_contents, lambda: part.contents):
        with pytest.raises(Docx4JException) as again:
            read()
        assert again.value is first.value


def test_a_change_of_the_parse_settings_reads_again(monkeypatch: pytest.MonkeyPatch):
    """D8: the record is keyed by the settings it was read under, as the private tree is (``_ParseSettings``)."""
    pkg = load(broken(1)[0])
    part = pkg.main_document_part
    with pytest.raises(Docx4JException) as first:
        _ = part.contents
    parses: list[Any] = []
    real = XmlPart._parse

    def counted(self: XmlPart, *args: Any, **kwargs: Any) -> Any:
        parses.append(self.part_name)
        return real(self, *args, **kwargs)

    monkeypatch.setattr(XmlPart, "_parse", counted)
    pkg.mc_load_mode = "word"
    with pytest.raises(Docx4JException) as second:
        _ = part.contents
    assert len(parses) == 1 and second.value is not first.value
    assert part.unreadable is second.value
    _assert_refused(second.value, "/word/document.xml", "Premature end of data in tag t line 2")


@pytest.mark.parametrize("number", sorted(CASES))
def test_every_save_writes_the_broken_part_as_read(number: int):
    """D3: an untouched save, a save after a read of another part, and after an edit of another part: byte for byte."""
    case = CASES[number]
    data, entries = broken(number)
    other = "word/header1.xml" if case.part == DOC else DOC

    pkg = load(data)
    with pytest.raises(Docx4JException):
        _ = _part(pkg, case.part).contents  # a failed read changes nothing either
    assert _unzip(_saved(pkg))[case.part] == entries[case.part]

    pkg = load(data)
    _part(pkg, other).read_contents()
    assert _unzip(_saved(pkg))[case.part] == entries[case.part]

    pkg = load(data)
    tree = _part(pkg, other).contents
    holder = tree.body.content if other == DOC else tree.content
    holder.append(wml(f"<w:p {W}><w:r><w:t>NEW</w:t></w:r></w:p>"))
    saved = _unzip(_saved(pkg))
    assert saved[case.part] == entries[case.part]
    assert b"NEW" in saved[other]


@pytest.mark.parametrize("number", [1, 6, 14, 15])
def test_a_replacement_is_saved(number: int):
    """``set_contents`` is the caller's decision (docx4j's ``setJaxbElement``): it clears the record, and is saved."""
    case = CASES[number]
    pkg = load(broken(number)[0])
    part = _part(pkg, case.part)
    with pytest.raises(Docx4JException):
        part.read_contents()
    replacement = load(_zip(BASE)).get_part("/" + case.part).contents
    part.set_contents(replacement)
    assert part.unreadable is None and not part.is_content_unreadable
    assert part.contents is replacement
    saved = load(_saved(pkg))
    reread = _part(saved, case.part).contents
    assert type(reread) is type(replacement)


def test_set_bytes_clears_the_record():
    pkg = load(broken(1)[0])
    part = pkg.main_document_part
    with pytest.raises(Docx4JException):
        _ = part.contents
    part.set_bytes(BASE[DOC])
    assert part.unreadable is None
    assert pkg.body.paragraphs[0].text == "FIRST PARAGRAPH"


@pytest.mark.parametrize("number", sorted(CASES))
def test_misplaced_reports_the_part_and_goes_on(number: int):
    """D9: ``part.not_well_formed`` with lxml's message, line and column; the walk goes on; nothing is changed."""
    case = CASES[number]
    data, entries = broken(number)
    pkg = load(data)
    findings = pkg.misplaced()
    assert [type(f) for f in findings] == [PartNotWellFormed]
    (finding,) = findings
    part = _part(pkg, case.part)
    syntax = syntax_error_of(part.unreadable)
    assert finding.code == "part.not_well_formed"
    assert finding.part_name == f"/{case.part}"
    assert finding.message == syntax.msg and finding.message.startswith(case.message)
    assert (finding.line, finding.column) == syntax.position
    assert finding.to_dict() == {
        "code": "part.not_well_formed",
        "part_name": f"/{case.part}",
        "line": syntax.position[0],
        "column": syntax.position[1],
        "message": syntax.msg,
    }
    # the part given alone: the same finding
    from docx4j_py.traversal import misplaced

    assert misplaced(part) == [finding]
    # every other part was read, and nothing changed
    assert _unzip(_saved(pkg)) == {**_unzip(_saved(load(data))), case.part: entries[case.part]}


def test_misplaced_goes_on_past_a_part_it_cannot_read():
    """Two parts not well formed: two findings, in the package's order, and the walk reads the parts between."""
    entries = dict(BASE)
    entries[DOC] = CASES[1].change(BASE[DOC])
    entries["word/header1.xml"] = CASES[14].change(BASE["word/header1.xml"])
    pkg = load(_zip(entries))
    findings = pkg.misplaced()
    assert [(f.code, f.part_name) for f in findings] == [
        ("part.not_well_formed", "/word/document.xml"),
        ("part.not_well_formed", "/word/header1.xml"),
    ]
    assert pkg.main_document_part.style_definitions_part._private_contents is not None  # read on the way


def test_a_cell_finding_still_counts_where_the_source_is_not_well_formed():
    """D9: a part replaced in code whose source is not well formed: ``source_tree()`` is None, so a finding in the
    replacement is the session's, and ``save(check=True)`` refuses it rather than raising the source's parse error."""
    from docx4j_py.model.content.errors import SaveCheckError

    pkg = load(broken(1)[0])
    part = pkg.main_document_part
    replacement = load(_zip(BASE)).main_document_part.contents
    replacement.body.content.insert(
        0, wml(f"<w:tbl {W}><w:tblGrid><w:gridCol w:w='100'/></w:tblGrid><w:tr><w:tc/></w:tr></w:tbl>")
    )
    part.set_contents(replacement)
    assert part.source_tree() is None
    with pytest.raises(SaveCheckError) as raised:
        _saved(pkg)
    assert raised.value.code == "cell.no_end_paragraph"


# ---------------------------------------------------------------------------
# the relationships and the content types (D5)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("strict", [False, True], ids=["lenient", "strict"])
@pytest.mark.parametrize("number", sorted(LOAD_CASES))
def test_a_relationships_part_not_well_formed_refuses_the_load(number: int, strict: bool):
    """D5, Q3: without its relationships the load cannot know the parts; before, case 17 loaded without the header and
    the numbering, and an untouched save dropped them. The content types are refused as before."""
    case, kind, prefix = LOAD_CASES[number]
    with pytest.raises(kind) as raised:
        load(broken(number)[0], options=LoadOptions(strict=strict))
    assert type(raised.value) is kind
    assert str(raised.value).startswith(prefix + case.message), str(raised.value)
    if kind is Docx4JException:
        _assert_refused(raised.value, prefix.removeprefix("Problem with part ").removesuffix(": "), case.message)


# ---------------------------------------------------------------------------
# the mc passes (D2)
# ---------------------------------------------------------------------------

_ALTERNATE = (
    f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
    '<mc:Choice Requires="w14"><w:p><w:r><w:t>CHOICE</w:t></w:r></w:p></mc:Choice>'
    "<mc:Fallback><w:p><w:r><w:t>FALLBACK</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>"
).encode()


def _with_alternate(data: bytes) -> bytes:
    """A block-level ``mc:AlternateContent`` first in the body."""
    return data.replace(b"<w:body>", b"<w:body>" + _ALTERNATE, 1)


OPTIONS = {
    "mce_preprocess": LoadOptions(mce_preprocess=True),
    "word": LoadOptions(mc_load_mode="word"),
    "preserve": LoadOptions(),
}


@pytest.mark.parametrize("alternate", [False, True], ids=["plain", "alternate-content"])
@pytest.mark.parametrize("route", sorted(OPTIONS))
def test_the_mc_passes_refuse_the_same(route: str, alternate: bool):
    """D2: the splice and the MCE pass inside ``_parse``'s ``try``: the same ``Docx4JException`` whichever parser meets
    the error, and Word mode refuses a part with or without ``mc:AlternateContent``. Before, lxml's own
    ``XMLSyntaxError`` came out of ``contents`` naming no part, and Word mode mended a part without the element."""
    entries = dict(BASE)
    source = _with_alternate(BASE[DOC]) if alternate else BASE[DOC]
    entries[DOC] = source
    whole = load(_zip(entries), options=OPTIONS[route])
    assert "FIRST PARAGRAPH" in whole.body.text  # well formed, it reads in this route
    entries[DOC] = CASES[1].change(source)
    pkg = load(_zip(entries), options=OPTIONS[route])
    with pytest.raises(Docx4JException) as raised:
        _ = pkg.main_document_part.contents
    _assert_refused(raised.value, "/word/document.xml", "Premature end of data in tag")
    assert pkg.main_document_part.is_content_unreadable


# ---------------------------------------------------------------------------
# a generic XML part (D3, Q9, Q10)
# ---------------------------------------------------------------------------


def test_a_custom_xml_part_not_well_formed_is_refused_when_read():
    """D3: ``tree`` raises naming the part, from lxml's error; the load registers it by its properties part's id (the
    sniff), and an untouched save writes it as read. The custom XML view's readers are left as they were (Q10)."""
    from docx4j_py.model.customxml import custom_xml_parts_of

    entries = _unzip((SAMPLES / "invoice2013.docx").read_bytes())
    item = entries["customXml/item1.xml"]
    entries["customXml/item1.xml"] = item[: len(item) // 2]
    data = _zip(entries)
    pkg = load(data)
    (part,) = [p for p in pkg.parts.parts() if isinstance(p, CustomXmlDataStoragePart)]
    assert part.item_id is not None and pkg.custom_xml_data_storage_parts[part.item_id] is part
    with pytest.raises(Docx4JException) as raised:
        _ = part.tree
    assert str(raised.value).startswith("Problem with part /customXml/item1.xml: ")
    assert isinstance(raised.value.__cause__, etree.XMLSyntaxError)
    assert str(raised.value) == f"Problem with part /customXml/item1.xml: {raised.value.__cause__.msg}"
    view = custom_xml_parts_of(pkg).get_item(part.item_id)
    assert view.document_element is None and view.get_xml() == ""  # Q10: as before
    assert _unzip(_saved(pkg))["customXml/item1.xml"] == entries["customXml/item1.xml"]


# ---------------------------------------------------------------------------
# docx4j's UnreadablePartTest, its five engine tests, over its own breaks
# ---------------------------------------------------------------------------

ORIGINAL = "ORIGINAL CONTENT stays"


def fixture(entry: str, old: str, new: str) -> bytes:
    """docx4j's ``fixture``: a document this package made and saved (a paragraph with ``w:spacing w:before="120"``, a
    header with the same paragraph, a default tab stop in the settings), with one replacement in `entry`'s XML."""
    pkg = create_package()
    mdp = pkg.main_document_part
    paragraph = f'<w:p {W}><w:pPr><w:spacing w:before="120"/></w:pPr><w:r><w:t>{ORIGINAL}</w:t></w:r></w:p>'
    mdp.contents.body.content.append(wml(paragraph))
    header = HeaderPart("/word/header.xml")
    rel = mdp.add_target_part(header)
    header.set_contents(wml(f"<w:hdr {W}>{paragraph}</w:hdr>"))
    mdp.contents.body.sect_pr.header_reference_or_footer_reference.append(
        wml(f'<w:headerReference {W} {R} w:type="default" r:id="{rel.id}"/>')
    )
    mdp.document_settings_part.contents.default_tab_stop = wml(f'<w:defaultTabStop {W} w:val="720"/>')
    entries = _unzip(_saved(pkg))
    xml = entries[entry].decode("utf-8")
    assert old in xml, f"{entry} holds {old}"
    entries[entry] = xml.replace(old, new).encode("utf-8")
    return _zip(entries)


def _after_save(pkg: Any, entry: str) -> str:
    return _unzip(_saved(pkg))[entry].decode("utf-8")


def test_a_main_part_that_cannot_be_read_is_not_replaced_by_an_edit_and_saves_its_bytes(
    monkeypatch: pytest.MonkeyPatch,
):
    """docx4j's ``aMainPartThatCannotBeReadIsNotReplacedByAnEditAndSavesItsBytes``."""
    pkg = load(fixture(DOC, "</w:body>", "</w:body"))
    mdp = pkg.main_document_part
    with pytest.raises(Docx4JException) as first:
        _ = mdp.contents
    assert mdp.is_content_unreadable
    monkeypatch.setattr(XmlPart, "_parse", lambda *a, **k: pytest.fail("a second attempt"))
    with pytest.raises(Docx4JException) as again:  # and again, without a second attempt
        _ = mdp.contents
    assert again.value is first.value
    # the edits are refused: there is nothing to edit
    with pytest.raises(Docx4JException):
        pkg.body.insert_paragraph("NEW")
    with pytest.raises(Docx4JException):
        mdp.contents.body.content.append(wml(f"<w:p {W}/>"))
    saved = _after_save(pkg, DOC)
    assert ORIGINAL in saved, "an untouched save keeps the source bytes"
    assert "</w:body" in saved and "</w:body>" not in saved
    assert "NEW" not in saved


def test_a_header_that_cannot_be_read_is_not_replaced():
    """docx4j's ``aHeaderThatCannotBeReadIsNotReplaced``."""
    pkg = load(fixture("word/header.xml", "</w:hdr>", "</w:hdr"))
    (header,) = [p for p in pkg.parts.parts() if isinstance(p, HeaderPart)]
    with pytest.raises(Docx4JException):
        _ = header.contents
    assert header.is_content_unreadable
    assert ORIGINAL in pkg.body.text, "the body still reads"
    saved = _after_save(pkg, "word/header.xml")
    assert "</w:hdr" in saved and "</w:hdr>" not in saved


def test_settings_that_cannot_be_read_are_not_replaced():
    """docx4j's ``settingsThatCannotBeReadAreNotReplaced``: the conveniences that edit them refused."""
    pkg = load(fixture("word/settings.xml", '<w:defaultTabStop w:val="720"/>', '<w:defaultTabStop w:val="720">'))
    settings = pkg.main_document_part.document_settings_part
    with pytest.raises(Docx4JException):
        _ = settings.contents
    assert settings.is_content_unreadable
    with pytest.raises(Docx4JException):
        settings.set_compat_setting("compatibilityMode", "15")
    with pytest.raises(Docx4JException):
        settings.set_default_compat_settings()
    saved = _after_save(pkg, "word/settings.xml")
    assert '<w:defaultTabStop w:val="720">' in saved and '<w:defaultTabStop w:val="720"/>' not in saved


def test_a_replacement_is_the_callers_decision():
    """docx4j's ``aReplacementIsTheCallersDecision``."""
    from docx4j_py.wml import el

    pkg = load(fixture(DOC, "</w:body>", "</w:body"))
    mdp = pkg.main_document_part
    with pytest.raises(Docx4JException):
        _ = mdp.contents
    replacement = el.document(body=el.body())
    mdp.set_contents(replacement)
    assert not mdp.is_content_unreadable
    assert mdp.contents is replacement
    pkg.body.insert_paragraph("NEW")
    assert "NEW" in _after_save(pkg, DOC)


def test_a_created_package_still_gets_new_content_from_the_conveniences():
    """docx4j's ``aCreatedPackageStillGetsNewContentFromTheConveniences``: a part with no content yet is not unreadable."""
    from docx4j_py.wml import el

    pkg = create_package()
    header = HeaderPart()
    pkg.main_document_part.add_target_part(header)
    assert not header.is_content_unreadable and header.unreadable is None
    with pytest.raises(Docx4JException, match="has no content"):
        _ = header.contents
    assert not header.is_content_unreadable  # no content yet is not a read that failed
    header.set_contents(el.hdr())
    header.contents.content.append(el.p())
    assert len(header.contents.content) == 1
    pkg.main_document_part.document_settings_part.set_compat_setting("compatibilityMode", "15")


# ---------------------------------------------------------------------------
# the sniffs (D4)
# ---------------------------------------------------------------------------


def _recovering_ns_map(data: bytes) -> dict[str, str]:
    """``_root_ns_map`` as it was, with lxml's recovery."""
    try:
        for _event, element in etree.iterparse(
            io.BytesIO(data), events=("start",), resolve_entities=False, huge_tree=True, recover=True
        ):
            return {p: u for p, u in element.nsmap.items() if p is not None}
    except etree.XMLSyntaxError:
        return {}
    return {}


def _corpus() -> list[Path]:
    paths = sorted(SAMPLES.glob("*.docx"))
    paths += sorted(p for p in (ROOT / "tests" / "fixtures").rglob("*") if p.suffix in (".docx", ".docm", ".dotx"))
    return paths


def test_the_sniffs_read_the_same_without_recovery_over_the_corpus():
    """D4: the root's declarations and a custom XML properties part's ``ds:itemID``, over every XML part of the corpus,
    are what the recovering sniffs gave (CR-031 section 2.5: no part of the corpus is not well formed)."""
    parts = props = 0
    for path in _corpus():
        for name, data in _unzip(path.read_bytes()).items():
            if not name.endswith((".xml", ".rels")):
                continue
            parts += 1
            assert _root_ns_map(data) == _recovering_ns_map(data), f"{path.name}:{name}"
        for part in load(path).parts.parts():
            if isinstance(part, CustomXmlDataStoragePropertiesPart):
                props += 1
                recovered = None
                for _event, element in etree.iterparse(
                    io.BytesIO(part.bytes_for_save), events=("start",), resolve_entities=False, recover=True
                ):
                    recovered = element.get("{http://schemas.openxmlformats.org/officeDocument/2006/customXml}itemID")
                    recovered = recovered or element.get("itemID")
                    break
                assert part.item_id == recovered, f"{path.name}:{part.part_name}"
    assert parts > 5000 and props > 40  # 5,671 and 44 at CR-031 Phase A


@pytest.mark.parametrize("number", sorted(CASES))
def test_the_root_sniff_reads_a_broken_part_as_far_as_its_root(number: int):
    """D4: lxml yields the root's start event before a later error, so only a broken root start tag gives ``{}``."""
    case = CASES[number]
    data = broken(number)[1][case.part]
    expected = {} if number == 13 else _root_ns_map(BASE[case.part])
    assert _root_ns_map(data) == expected
    assert expected or number == 13


def test_a_properties_part_cut_after_its_root_still_gives_its_id():
    """Section 2.4: the custom XML properties part cut after its root's start tag still gives its ``ds:itemID``."""
    entries = _unzip((SAMPLES / "invoice2013.docx").read_bytes())
    whole = load(_zip(entries))
    (props,) = [p for p in whole.parts.parts() if isinstance(p, CustomXmlDataStoragePropertiesPart)]
    expected = props.item_id
    xml = entries["customXml/itemProps1.xml"]
    root_end = re.search(rb"<ds:datastoreItem[^>]*>", xml).end()
    entries["customXml/itemProps1.xml"] = xml[: root_end + 5]
    pkg = load(_zip(entries))
    (props,) = [p for p in pkg.parts.parts() if isinstance(p, CustomXmlDataStoragePropertiesPart)]
    assert expected and props.item_id == expected


# ---------------------------------------------------------------------------
# the census (D4): no parser in the package recovers
# ---------------------------------------------------------------------------


def _package_sources() -> list[tuple[Path, str]]:
    import docx4j_py

    package = Path(docx4j_py.__file__).parent
    return [
        (path, f"docx4j_py/{path.relative_to(package).as_posix()}")
        for path in sorted(package.rglob("*.py"))
        if "__pycache__" not in path.parts
    ]


def test_no_parser_in_the_package_recovers():
    """Section 4.4's census: no ``recover=True`` (lxml's default is strict), no HTML parser (which recovers by default),
    and the fork's config with ``recover`` False."""
    from docx4j_py.runtime import RECOVER, parser, parser_config

    offenders = []
    for path, relative in _package_sources():
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.keyword) and node.arg == "recover":
                if not (isinstance(node.value, ast.Constant) and node.value.value is False) and not (
                    isinstance(node.value, ast.Name) and node.value.id == "RECOVER"
                ):
                    offenders.append(f"{relative}:{node.value.lineno}: recover={ast.unparse(node.value)}")
            elif isinstance(node, ast.Call):
                name = node.func.attr if isinstance(node.func, ast.Attribute) else getattr(node.func, "id", None)
                if name in ("HTMLParser", "HTMLPullParser"):
                    offenders.append(f"{relative}:{node.lineno}: {name}(...)")
            elif isinstance(node, (ast.Import, ast.ImportFrom)):
                module = node.module if isinstance(node, ast.ImportFrom) else None
                names = [module] if module else [alias.name for alias in node.names]
                if any(n and (n == "lxml.html" or n.startswith("lxml.html.")) for n in names):
                    offenders.append(f"{relative}:{node.lineno}: imports lxml.html")
    assert not offenders, "\n".join(offenders)
    assert RECOVER is False
    assert parser_config().recover is False
    assert parser().config.recover is False


def test_the_engine_s_own_lxml_parsers_refuse_xml_not_well_formed():
    """Section 4.4: the module-level parsers (CR-030's pass, the MCE pass, the splice, a generic part, the fragments)
    are strict, as lxml's default is: none mends a cut document."""
    from docx4j_py import fragments
    from docx4j_py.openpackaging import breaks, mc_load_mode, mce
    from docx4j_py.openpackaging.parts import default_xml_part

    cut = f"<w:document {W}><w:body><w:p><w:r><w:t>SEC".encode()
    for parser in (breaks._PARSER, mce._PARSER, mc_load_mode._PARSER, default_xml_part._PARSER, fragments._parser):
        with pytest.raises(etree.XMLSyntaxError):
            etree.fromstring(cut, parser)


# ---------------------------------------------------------------------------
# downstream of the refusal (section 2.3): the text model, the resolver, the facade
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("number", [n for n in sorted(CASES) if CASES[n].part == DOC])
def test_the_text_model_names_the_main_part_and_lxml_s_reason(number: int):
    """Cases 1 to 13: CR-027 section 4.6's refusal, its reason lxml's (D7), where the recovering parse read on."""
    from docx4j_py.model.text import extract

    pkg = load(broken(number)[0])
    with pytest.raises(Docx4JException) as raised:
        extract(pkg)
    syntax = syntax_error_of(raised.value)
    assert str(raised.value) == (
        f"The content of /word/document.xml could not be read, so it cannot be read as text: {syntax.msg}"
    )
    assert str(syntax.msg).startswith(CASES[number].message)


def test_the_text_model_reads_the_body_without_a_header_it_cannot_read():
    """Case 14: ``content.not_read`` for the header, the body read (before, ``HEA`` with no warning)."""
    from docx4j_py.model.text import TextWarning, extract, plain_text

    result = extract(load(broken(14)[0]), include="headers")
    assert [(w.kind, w.address, w.part_name) for w in result.warnings] == [
        (TextWarning.CONTENT_NOT_READ, "header:rId5", "/word/header1.xml")
    ]
    assert "FIRST PARAGRAPH" in plain_text(result) and "HEA" not in plain_text(result)


def test_the_content_api_refuses_unreadable_styles_and_reads_without_numbering():
    """D6: case 15's resolver and ``body.text`` raise the styles part's exception; case 16's ``body.text`` reads."""
    pkg = load(broken(15)[0])
    styles = pkg.main_document_part.style_definitions_part
    with pytest.raises(Docx4JException) as raised:
        pkg.get_property_resolver()
    _assert_refused(raised.value, "/word/styles.xml", CASES[15].message)
    with pytest.raises(Docx4JException) as again:
        _ = pkg.body.text
    assert again.value is raised.value is styles.unreadable
    pkg = load(broken(16)[0])
    assert "SECOND PARAGRAPH" in pkg.body.text


def test_the_python_docx_facade_refuses_too():
    """Section 2.2: ``docx4j_py.docx.Document`` of case 1 gave ``FIRST PARAGRAPH``, ``SEC``; now the read is refused."""
    from docx4j_py.docx import Document

    with pytest.raises(Docx4JException, match="^Problem with part /word/document.xml: Premature end of data"):
        _ = [p.text for p in Document(io.BytesIO(broken(1)[0])).paragraphs]


# --- a relationships part in a wrong namespace -------------------------------------------


@pytest.mark.parametrize("entry", ["word/_rels/document.xml.rels", "_rels/.rels"])
def test_a_relationships_part_in_a_wrong_namespace_is_refused(entry: str) -> None:
    """It would parse as no relationships with nothing said, and a save would lose every part it named."""
    entries = _base()
    entries[entry] = entries[entry].replace(b"http://schemas.openxmlformats.org/package/2006/relationships", b"urn:wrong")
    with pytest.raises(Docx4JException, match="not {http://schemas.openxmlformats.org/package/2006/relationships}Relationships"):
        load(_zip(entries))


def test_a_relationships_part_with_another_root_is_refused() -> None:
    entries = _base()
    entries["word/_rels/document.xml.rels"] = b'<Other xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    with pytest.raises(Docx4JException, match="/word/_rels/document.xml.rels: its root is"):
        load(_zip(entries))


@pytest.mark.parametrize("entry", ["word/_rels/document.xml.rels", "_rels/.rels"])
@pytest.mark.parametrize("count", [1, -1], ids=["one", "every"])
def test_a_relationship_in_a_wrong_namespace_is_refused(entry: str, count: int) -> None:
    """``rn-06``, ``rn-07``: the root is right, so only the children say, and a save would lose what each named."""
    entries = _base()
    held = entries[entry].replace(b"<Relationships ", b'<Relationships xmlns:x="urn:wrong" ', 1)
    entries[entry] = held.replace(b"<Relationship ", b"<x:Relationship ", count)
    with pytest.raises(Docx4JException, match="it holds {urn:wrong}Relationship, in a wrong namespace"):
        load(_zip(entries))


def test_a_foreign_element_beside_the_relationships_is_ignored() -> None:
    """It names no part, so nothing is lost, as docx4j reads it."""
    entries = _base()
    entries["_rels/.rels"] = entries["_rels/.rels"].replace(b"</Relationships>", b'<Other xmlns="urn:x"/></Relationships>')
    assert load(_zip(entries)).main_document_part is not None


def test_an_empty_relationships_part_in_the_right_namespace_loads() -> None:
    """Nothing is lost where it names nothing."""
    entries = _base()
    entries["word/_rels/document.xml.rels"] = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    assert load(_zip(entries)).main_document_part is not None
