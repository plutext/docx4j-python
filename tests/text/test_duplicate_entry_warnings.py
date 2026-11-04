"""The text model and a package holding two entries of one name: ``package.duplicate_entry``. CR-032 D8 and D9.

WD REQ-206 and REQ-1001. The load decides (``LoadOptions(duplicate_entries=)``),
and a reading of the package it loaded does not decide again:

- under ``first``, :func:`~docx4j_py.model.text.extract` carries each of the
  package's :class:`~docx4j_py.openpackaging.load.DuplicateEntry` records into
  ``TextResult.warnings``, first, with no address, the selector the entry's name
  as the first entry spells it and ``detail`` the number of entries;
  ``TextWarning.to_dict`` writes ``detail`` only where it is set (Q7);
- ``TextOptions`` accepts ``duplicate_entries`` (``refuse``, ``first``) by name
  (Q5), last in section 10's order, and the adapter passes it to the load;
- an embedded document or an ``altChunk`` holding a duplicate is refused by its
  own load, under the default: ``content.not_read`` and
  ``altchunk.not_expanded`` (Q6).
"""

from __future__ import annotations

import io
import warnings
import zipfile

import pytest
from adapter import BEYOND_DOCX4J, TextsAdapter
from mini_docx import MiniDocx

from docx4j_py import create_package, load
from docx4j_py.fragments import wml
from docx4j_py.model.text import (
    OPTION_NAMES,
    TextOptionError,
    TextOptions,
    TextWarning,
    extract,
    plain_text,
)
from docx4j_py.openpackaging import DuplicateEntryException
from docx4j_py.openpackaging.load import LoadOptions
from docx4j_py.openpackaging.parts.binary_part import EmbeddedPackagePart

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
DOC = "word/document.xml"
DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
FIRST = LoadOptions(duplicate_entries="first")


def _doubled(text: str = "BODY") -> bytes:
    """A document made here whose ``word/document.xml`` is followed by a copy: ``<text> FIRST``, then SECOND."""
    pkg = create_package()
    pkg.main_document_part.contents.body.content.append(wml(f"<w:p {W}><w:r><w:t>{text} MARKER</w:t></w:r></w:p>"))
    out = io.BytesIO()
    pkg.save(out)
    with zipfile.ZipFile(io.BytesIO(out.getvalue())) as z:
        entries = [(info.filename, z.read(info)) for info in z.infolist()]
    buf = io.BytesIO()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # zipfile's "Duplicate name"
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in entries:
                if name == DOC:
                    z.writestr(name, data.replace(b"MARKER", b"FIRST"))
                    z.writestr("Word/Document.xml", data.replace(b"MARKER", b"SECOND"))
                else:
                    z.writestr(name, data)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# the warning in the result
# ---------------------------------------------------------------------------


def test_extract_under_first_carries_the_warning_with_its_count():
    pkg = load(_doubled(), options=FIRST)
    result = extract(pkg)
    assert plain_text(result) == "BODY FIRST\n"
    (warning,) = result.warnings
    assert warning.kind == TextWarning.PACKAGE_DUPLICATE_ENTRY == "package.duplicate_entry"
    assert (warning.address, warning.selector, warning.detail) == (None, DOC, 2)
    assert warning.to_dict() == {"kind": "package.duplicate_entry", "address": None, "selector": DOC, "detail": 2}
    assert str(warning) == f"package.duplicate_entry ({DOC}) (2)"


def test_the_warning_comes_first_and_once_in_every_view():
    """A reading of the package raises it once, ahead of the reading's own (here, a field with no result)."""
    pkg = load(_doubled(), options=FIRST)
    pkg.main_document_part.contents.body.content.append(
        wml(f'<w:p {W}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    )
    for view in ("accepted", "original", "markup"):
        kinds = [w.kind for w in extract(pkg, view=view).warnings]
        assert kinds[0] == TextWarning.PACKAGE_DUPLICATE_ENTRY
        assert kinds.count(TextWarning.PACKAGE_DUPLICATE_ENTRY) == 1


def test_a_warning_without_a_count_writes_no_detail():
    warning = TextWarning(TextWarning.PART_NOT_READ, None, part_name="/word/styles.xml")
    assert warning.detail is None
    assert warning.to_dict() == {"kind": "part.not_read", "address": None, "selector": "/word/styles.xml"}


def test_a_reading_does_not_decide_again():
    """``duplicate_entries`` is the load's: a reading's value changes nothing, either way."""
    pkg = load(_doubled(), options=FIRST)
    result = extract(pkg, duplicate_entries="refuse")
    assert [w.kind for w in result.warnings] == [TextWarning.PACKAGE_DUPLICATE_ENTRY]
    assert plain_text(result) == "BODY FIRST\n"
    clean = create_package()
    assert extract(clean, duplicate_entries="first").warnings == []
    with pytest.raises(DuplicateEntryException):
        load(_doubled())  # the default load refuses, so a reading never sees one


# ---------------------------------------------------------------------------
# the option by name (REQ-1001)
# ---------------------------------------------------------------------------


def test_the_option_is_accepted_by_name_with_its_two_values():
    assert TextOptions().duplicate_entries == "refuse"
    assert TextOptions(duplicate_entries="first").duplicate_entries == "first"
    assert TextOptions.of({"duplicate_entries": "FIRST"}).duplicate_entries == "first"
    assert OPTION_NAMES[-1] == "duplicate_entries"
    assert list(TextOptions().to_dict())[-1] == "duplicate_entries"
    assert TextOptions().to_dict()["duplicate_entries"] == "refuse"
    assert repr(TextOptions(duplicate_entries="first")) == "TextOptions(duplicate_entries='first')"
    assert TextOptions(duplicate_entries="first").copy().duplicate_entries == "first"


@pytest.mark.parametrize("value", ["last", "", "true"])
def test_another_value_is_refused(value):
    with pytest.raises(TextOptionError) as raised:
        TextOptions(duplicate_entries=value)
    assert (raised.value.option, raised.value.statement, raised.value.reason) == (
        "duplicate_entries",
        "REQ-206",
        "invalid",
    )


def test_the_adapter_passes_the_option_to_the_load(tmp_path):
    fixture = tmp_path / "dup.docx"
    fixture.write_bytes(_doubled())
    adapter = TextsAdapter()
    assert "REQ-206" in adapter.implemented() and "REQ-206" in BEYOND_DOCX4J
    output = adapter.extract(fixture, "accepted", "duplicate_entries: first", "txt")
    assert output.text == "BODY FIRST\n"
    assert [(k, v) for k, v in output.findings.items() if k.startswith("warning")] == [
        (f"warning\tpackage.duplicate_entry\t{DOC}", "1")
    ]
    with pytest.raises(DuplicateEntryException):
        adapter.extract(fixture, "accepted", "default", "txt")


# ---------------------------------------------------------------------------
# packages inside packages (D9, Q6)
# ---------------------------------------------------------------------------


def test_an_embedded_document_holding_a_duplicate_is_content_not_read():
    outer = create_package()
    mdp = outer.main_document_part
    embedded = EmbeddedPackagePart("/word/embeddings/dup.docx", DOCX)
    embedded.set_bytes(_doubled("EMBEDDED"))
    rel_id = mdp.add_target_part(embedded).id
    mdp.contents.body.content.append(
        wml(
            f'<w:p {W} xmlns:o="urn:schemas-microsoft-com:office:office" {R}><w:r><w:object>'
            f'<o:OLEObject Type="Embed" ProgID="Word.Document.12" r:id="{rel_id}"/></w:object></w:r></w:p>'
        )
    )
    result = extract(outer, embedded=True)
    # at the object's paragraph: the package did not load, so there is no nested reading to raise it at its root
    assert [(w.kind, w.address, w.part_name) for w in result.warnings] == [
        (TextWarning.CONTENT_NOT_READ, "body/0", "/word/embeddings/dup.docx")
    ]
    assert "EMBEDDED" not in plain_text(result)


def test_a_chunk_holding_a_duplicate_is_not_expanded():
    pkg = (
        MiniDocx.body("<w:p><w:r><w:t>OUTER</w:t></w:r></w:p><w:altChunk r:id='rIdC'/>")
        .part("chunk.docx", f"{_CT}.document.main+xml", f"{_REL}/aFChunk", "rIdC", _doubled("CHUNK"))
        .load()
    )
    result = extract(pkg)
    assert [(w.kind, w.address, w.part_name) for w in result.warnings] == [
        (TextWarning.ALTCHUNK_NOT_EXPANDED, "body/1", "/word/chunk.docx")
    ]
    assert plain_text(result) == "OUTER\n"
