"""A part the text model cannot read, named: docx4j ``4fdd80063``'s text-model half. CR-027 sections 2.4 and 4.6.

docx4j's ``UnreadablePartTest`` at ``4fdd80063`` (``VERSION_17_3_6``), its four
text-model tests, over a document this package made and saved, with one
replacement in the named part's XML (docx4j's ``fixture``):

- ``theTextModelSaysWhichPartCouldNotBeReadAndWhy``: the package's text, and
  the main part's, refused with ``Docx4JException`` naming the part and why,
  the part's own exception its ``__cause__``;
- ``theTextModelOfAHeaderSaysWhichPartCouldNotBeRead``: a header's, the same;
- ``anEmbeddedDocumentThatCannotBeReadIsContentNotRead``: one level down,
  ``content.not_read`` naming the embedded package's part (REQ-512);
- ``unreadableHeadersAndNotesAreContentNotRead``: a header under ``include:
  headers`` and the footnotes, ``content.not_read`` with the body read (REQ-801).

Then every other story part (the endnotes, the comments, a footer, the
glossary) as docx4j's ``storyContent`` has it, a chunk whose main part cannot
be read, and the part's route leaving a part as it was (REQ-204).

**The break** is docx4j's: a closing tag cut short (``</w:body>`` to
``</w:body`` in the main part, the root's own elsewhere, ``</w:hdr>`` to
``</w:hdr``), and the reason lxml's for it (``expected '>'``, at the cut's line
and column). Until CR-031 this package's parse recovered from that (lxml's
``recover``, under the xsdata fork's handler) and read the part, so the tests
broke the root's start tag instead (``<w:document`` to ``<<w:document``), with
the parser's reason "Failed to create target class"; since the fork's stage 15
nothing recovers (CR-031 section 4.7), which retires CR-027 section 10.8's
departure 1.
"""

from __future__ import annotations

import io
import zipfile
from typing import NamedTuple

import pytest
from mini_docx import NS, MiniDocx

from docx4j_py import create_package, load
from docx4j_py.fragments import wml
from docx4j_py.model.text import TextWarning, extract, plain_text
from docx4j_py.openpackaging.exceptions import Docx4JException
from docx4j_py.openpackaging.parts.binary_part import EmbeddedPackagePart
from docx4j_py.openpackaging.parts.wml import FootnotesPart, HeaderPart

ORIGINAL = "ORIGINAL CONTENT stays"
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
PARAGRAPH = f'<w:p {W}><w:pPr><w:spacing w:before="120"/></w:pPr><w:r><w:t>{ORIGINAL}</w:t></w:r></w:p>'
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _unzip(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {name: z.read(name) for name in z.namelist()}


def _zip(entries: dict[str, bytes]) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)
    return out.getvalue()


def _cut(xml: bytes, part: str) -> int:
    """Where docx4j's break goes: after ``</w:body`` in the main part, after the root's closing tag's name elsewhere."""
    if part.endswith("word/document.xml") and b"</w:body>" in xml:
        return xml.index(b"</w:body>") + len(b"</w:body")
    return xml.rindex(b">")


def _break(data: bytes, part: str) -> bytes:
    """The package with docx4j's break in `part`: a closing tag cut short, ``</w:body>`` to ``</w:body`` in the main
    part and the root's (``</w:hdr>`` to ``</w:hdr``) elsewhere."""
    entries = _unzip(data)
    xml = entries[part]
    at = _cut(xml, part)
    assert xml[at : at + 1] == b">"
    entries[part] = xml[:at] + xml[at + 1 :]
    return _zip(entries)


def _expected_reason(data: bytes, part: str) -> str:
    """lxml's reason for docx4j's break in `data`: ``expected '>'`` at the first byte after the cut, 1-based."""
    xml = _unzip(data)[part]
    if part.endswith("word/document.xml") and b"</w:body<" in xml:
        at = xml.index(b"</w:body<") + len(b"</w:body")
    else:
        assert not xml.endswith(b">")  # the root's closing tag, cut short at the end
        at = len(xml)
    line = xml.count(b"\n", 0, at) + 1
    column = at - (xml.rfind(b"\n", 0, at) + 1) + 1
    return f"expected '>', line {line}, column {column}"


def _saved(pkg) -> bytes:
    out = io.BytesIO()
    pkg.save(out)
    return out.getvalue()


def _made() -> bytes:
    """docx4j's ``fixture`` before the break: a document this package made and saved, a paragraph and a header with
    the same paragraph."""
    pkg = create_package()
    mdp = pkg.main_document_part
    mdp.contents.body.content.append(wml(PARAGRAPH))
    header = HeaderPart()
    rel = mdp.add_target_part(header)
    header.set_contents(wml(f"<w:hdr {W}>{PARAGRAPH}</w:hdr>"))
    reference = wml(f'<w:headerReference {W} {R} w:type="default" r:id="{rel.id}"/>')
    mdp.contents.body.sect_pr.header_reference_or_footer_reference.append(reference)
    return _saved(pkg)


def fixture(part: str) -> bytes:
    """docx4j's ``fixture``: :func:`_made`, with `part`'s XML broken as docx4j breaks it."""
    return _break(_made(), part)


def _header(pkg) -> HeaderPart:
    (header,) = [p for p in pkg.parts.parts() if isinstance(p, HeaderPart)]
    return header


def _assert_names_the_part(read, part_name: str, reason: str) -> None:
    """docx4j's ``assertNamesThePart``: refused, naming the part and why, the part's own exception the cause."""
    with pytest.raises(Docx4JException) as raised:
        read()
    # the reason: the deepest cause's message, on one line: lxml's (its msg, CR-031 section 4.7)
    assert str(raised.value) == f"The content of {part_name} could not be read, so it cannot be read as text: {reason}"
    # the part's own exception, which names it, is the cause, and gives the same reason
    cause = raised.value.__cause__
    assert isinstance(cause, Docx4JException) and str(cause) == f"Problem with part {part_name}: {reason}"


def _assert_not_read(result, part_name: str) -> None:
    """docx4j's ``assertNotRead``: the body read, and ``content.not_read`` naming the part."""
    assert any(ORIGINAL in block.text for block in result.blocks), "the body is still read"
    found = [w for w in result.warnings if w.kind == TextWarning.CONTENT_NOT_READ and w.part_name == part_name]
    assert found, result.warnings


# ---------------------------------------------------------------------------
# UnreadablePartTest's four
# ---------------------------------------------------------------------------


def test_the_text_model_says_which_part_could_not_be_read_and_why():
    """docx4j's ``theTextModelSaysWhichPartCouldNotBeReadAndWhy``: the package's text and the main part's."""
    data = fixture("word/document.xml")
    reason = _expected_reason(data, "word/document.xml")
    assert reason.startswith("expected '>', line 2, column ")
    pkg = load(data)
    _assert_names_the_part(lambda: extract(pkg), "/word/document.xml", reason)  # before CR-027, empty
    _assert_names_the_part(lambda: extract(pkg.main_document_part), "/word/document.xml", reason)


def test_the_text_model_of_a_header_says_which_part_could_not_be_read():
    """docx4j's ``theTextModelOfAHeaderSaysWhichPartCouldNotBeRead``; before CR-027, ``TypeError("not a story part")``."""
    data = fixture("word/header1.xml")
    pkg = load(data)
    reason = _expected_reason(data, "word/header1.xml")
    _assert_names_the_part(lambda: extract(_header(pkg)), "/word/header1.xml", reason)


def test_an_embedded_document_that_cannot_be_read_is_content_not_read():
    """docx4j's ``anEmbeddedDocumentThatCannotBeReadIsContentNotRead``: one level down, REQ-512, at the nested root."""
    outer = create_package()
    mdp = outer.main_document_part
    embedded = EmbeddedPackagePart("/word/embeddings/unreadable.docx", DOCX)
    embedded.set_bytes(fixture("word/document.xml"))
    rel_id = mdp.add_target_part(embedded).id
    mdp.contents.body.content.append(
        wml(
            f'<w:p {W} xmlns:o="urn:schemas-microsoft-com:office:office" {R}><w:r><w:object>'
            f'<o:OLEObject Type="Embed" ProgID="Word.Document.12" r:id="{rel_id}"/></w:object></w:r></w:p>'
        )
    )
    result = extract(outer, embedded=True)
    found = [w for w in result.warnings if w.kind == TextWarning.CONTENT_NOT_READ]
    # named for the embedded package's part, the one the host package has (the inner one is /word/document.xml)
    assert [(w.address, w.part_name) for w in found] == [("body/0/object:0", "/word/embeddings/unreadable.docx")]
    assert not [b for b in result.blocks if b.story == "EMBEDDED"]


def test_unreadable_headers_and_notes_are_content_not_read():
    """docx4j's ``unreadableHeadersAndNotesAreContentNotRead``: REQ-801, the body still read. Before CR-027 both
    were passed over silently (docx4j's notes threw a NullPointerException)."""
    pkg = load(fixture("word/header1.xml"))
    _assert_not_read(extract(pkg, include="headers"), "/word/header1.xml")  # headers are read on request

    with_note = create_package()
    mdp = with_note.main_document_part
    notes = FootnotesPart()
    mdp.add_target_part(notes)
    notes.set_contents(
        wml(f'<w:footnotes {W}><w:footnote w:id="1"><w:p><w:r><w:t>NOTE</w:t></w:r></w:p></w:footnote></w:footnotes>')
    )
    mdp.contents.body.content.append(
        wml(f'<w:p {W}><w:r><w:t>{ORIGINAL}</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>')
    )
    saved = _saved(with_note)
    assert "NOTE" in plain_text(extract(load(saved)))
    _assert_not_read(extract(load(_break(saved, "word/footnotes.xml"))), "/word/footnotes.xml")


# ---------------------------------------------------------------------------
# every story part, and a chunk
# ---------------------------------------------------------------------------

_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"


class Story(NamedTuple):
    """A story part that is not the main part, and how docx4j's ``storyContent`` reports it."""

    content_type: str
    rel_type: str
    rel_id: str
    xml: str
    #: The text it holds, and the options that read it.
    text: str
    options: dict
    #: Where its ``content.not_read`` is: none for the notes, the comments and the glossary, a header's prefix.
    address: str | None


def _p(text: str) -> str:
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


STORIES = {
    "footnotes.xml": Story(
        f"{_CT}.footnotes+xml",
        f"{_REL}/footnotes",
        "rIdFn",
        f'<w:footnotes {NS}><w:footnote w:id="1">{_p("FN_TEXT")}</w:footnote></w:footnotes>',
        "FN_TEXT",
        {},
        None,
    ),
    "endnotes.xml": Story(
        f"{_CT}.endnotes+xml",
        f"{_REL}/endnotes",
        "rIdEn",
        f'<w:endnotes {NS}><w:endnote w:id="1">{_p("EN_TEXT")}</w:endnote></w:endnotes>',
        "EN_TEXT",
        {},
        None,
    ),
    "comments.xml": Story(
        f"{_CT}.comments+xml",
        f"{_REL}/comments",
        "rIdCm",
        f'<w:comments {NS}><w:comment w:id="0" w:author="A">{_p("CM_TEXT")}</w:comment></w:comments>',
        "CM_TEXT",
        {"comments": True},
        None,
    ),
    "header1.xml": Story(
        f"{_CT}.header+xml",
        f"{_REL}/header",
        "rIdH",
        f"<w:hdr {NS}>{_p('HD_TEXT')}</w:hdr>",
        "HD_TEXT",
        {"include": "headers"},
        "header:rIdH",
    ),
    "footer1.xml": Story(
        f"{_CT}.footer+xml",
        f"{_REL}/footer",
        "rIdF",
        f"<w:ftr {NS}>{_p('FT_TEXT')}</w:ftr>",
        "FT_TEXT",
        {"include": "footers"},
        "footer:rIdF",
    ),
    "glossary/document.xml": Story(
        f"{_CT}.document.glossary+xml",
        f"{_REL}/glossaryDocument",
        "rIdG",
        f'<w:glossaryDocument {NS}><w:docParts><w:docPart><w:docPartPr><w:name w:val="G"/></w:docPartPr>'
        f"<w:docPartBody>{_p('GL_TEXT')}</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>",
        "GL_TEXT",
        {"include": "glossary"},
        None,
    ),
}
BODY = (
    f'<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>{ORIGINAL}</w:t></w:r><w:commentRangeEnd w:id="0"/>'
    '<w:r><w:commentReference w:id="0"/></w:r><w:r><w:footnoteReference w:id="1"/></w:r>'
    '<w:r><w:endnoteReference w:id="1"/></w:r></w:p>'
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>'
    "</w:sectPr>"
)


def _stories(broken: str | None = None) -> bytes:
    made = MiniDocx.body(BODY)
    for name, story in STORIES.items():
        made.part(name, story.content_type, story.rel_type, story.rel_id, story.xml)
    data = made.save()
    return data if broken is None else _break(data, f"word/{broken}")


def _everything(result) -> str:
    """The plain text and the comments' text."""
    return plain_text(result) + "".join(block.text for comment in result.comments.values() for block in comment.blocks)


@pytest.mark.parametrize("name", STORIES)
def test_every_other_story_part_that_cannot_be_read_is_content_not_read(name):
    """CR-027 section 4.6: each story part but the main one, once, at docx4j's address; the rest is read."""
    story = STORIES[name]
    readable = extract(load(_stories()), **story.options)
    assert story.text in _everything(readable)
    assert not [w for w in readable.warnings if w.kind == TextWarning.CONTENT_NOT_READ]
    result = extract(load(_stories(name)), **story.options)
    found = [w for w in result.warnings if w.kind == TextWarning.CONTENT_NOT_READ]
    assert [(w.address, w.part_name) for w in found] == [(story.address, f"/word/{name}")]
    assert ORIGINAL in plain_text(result)
    assert story.text not in _everything(result)
    # not asked for, it is not read, and nothing is said
    if story.options:
        assert not [w for w in extract(load(_stories(name))).warnings if w.kind == TextWarning.CONTENT_NOT_READ]


def test_a_chunk_whose_main_part_cannot_be_read_is_content_not_read():
    """A WordprocessingML ``altChunk`` one level down, as an embedded document: ``content.not_read`` at the chunk's
    root, naming the chunk, and nothing else of it; no ``altchunk.not_expanded``, the chunk was read."""
    pkg = (
        MiniDocx.body(f"{_p(ORIGINAL)}<w:altChunk r:id='rIdC'/>")
        .part("chunk.docx", f"{_CT}.document.main+xml", f"{_REL}/aFChunk", "rIdC", fixture("word/document.xml"))
        .load()
    )
    result = extract(pkg)
    assert [(w.kind, w.address, w.part_name) for w in result.warnings] == [
        (TextWarning.CONTENT_NOT_READ, "body/1/altchunk:0", "/word/chunk.docx")
    ]
    assert plain_text(result) == f"{ORIGINAL}\n"


def test_a_header_of_a_package_whose_main_part_cannot_be_read_is_read():
    """A header read as a part (docx4j's ``Texts.of(Part)``) reads that part alone, the main part unread."""
    pkg = load(fixture("word/document.xml"))
    assert plain_text(extract(_header(pkg))) == f"{ORIGINAL}\n"


def test_a_part_read_as_a_part_is_left_as_it_was():
    """REQ-204 on the part's route: telling a story part from content no longer unmarshals it (``contents``)."""
    pkg = load(_stories())
    assert plain_text(extract(pkg.main_document_part)) == f"{ORIGINAL}[1][1]\n"  # the part's route: no notes
    assert plain_text(extract(_header(pkg))) == "HD_TEXT\n"
    assert not pkg.main_document_part.is_unmarshalled
    assert not _header(pkg).is_unmarshalled
