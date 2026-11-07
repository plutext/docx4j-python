"""A prefix ``mc:Ignorable`` names with no declaration on the part's root. WD REQ-323 (0.103); CR-040 section 9.

Read as though the list didn't hold it, ``mce.undeclared_prefix`` once a part, on ``part.warnings`` and in the text
model. A save that writes the part anew drops the prefix from the list, unless what the part holds now declares it;
one that writes the source bytes keeps it. The oracle is the ``undeclared-ignorable`` round
(``tests/fixtures/undeclared-ignorable/``): Word repairs each file, one prompt an open, and loses nothing, but
``ui-05``, which it refuses.
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from pathlib import Path
from typing import Any

import pytest
from lxml import etree

from docx4j_py import load
from docx4j_py.model.text import TextWarning, extract, plain_text
from docx4j_py.openpackaging.mce import UndeclaredPrefix
from docx4j_py.openpackaging.parts.xml_part import XmlPart

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "undeclared-ignorable"
IGNORABLE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"

#: Each file Word repairs, and the parts whose root lists an undeclared prefix, with the prefixes.
REPAIRED = {
    "ui-01-main-unknown": {"/word/document.xml": ("zz9",)},
    "ui-02-main-known": {"/word/document.xml": ("w14",)},
    "ui-03-main-several": {"/word/document.xml": ("w15", "wp14")},
    "ui-04-main-lower-known": {"/word/document.xml": ("w14",)},
    "ui-10-header": {"/word/header1.xml": ("zz9",)},
    "ui-11-footnotes": {"/word/footnotes.xml": ("zz9",)},
    "ui-12-styles": {"/word/styles.xml": ("zz9",)},
    "ui-13-settings": {"/word/settings.xml": ("zz9",)},
    "ui-14-all-parts-known": {
        name: ("w14",) for name in ("/word/document.xml", "/word/styles.xml", "/word/settings.xml", "/word/footnotes.xml", "/word/header1.xml")
    },
}
#: Silent in Word: the control, the prefix declared on the root, a leading space in the list.
SILENT = ["ui-00-control", "ui-06-main-declared-unknown", "ui-07-main-leading-space"]


def source(name: str) -> bytes:
    return (FIXTURES / f"{name}.docx").read_bytes()


def word_text(name: str) -> str:
    return (FIXTURES / f"{name}-word365.txt").read_text(encoding="utf-8")


def every_part_read(package: Any) -> None:
    """Read each part's contents, so a save writes every one anew."""
    for part in package.parts.parts():
        if isinstance(part, XmlPart) and part.model_class is not None:
            part.contents


def saved(package: Any) -> bytes:
    out = io.BytesIO()
    package.save(out)
    return out.getvalue()


def undeclared(data: bytes) -> dict[str, list[str]]:
    """Each part of a package whose root's ``mc:Ignorable`` names a prefix that root doesn't declare."""
    found: dict[str, list[str]] = {}
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for name in z.namelist():
            if not name.endswith(".xml") or name.startswith("["):
                continue
            root = etree.fromstring(z.read(name))
            missing = [prefix for prefix in (root.get(IGNORABLE) or "").split() if prefix not in root.nsmap]
            if missing:
                found[f"/{name}"] = missing
    return found


def test_every_fixture_is_the_file_word_read() -> None:
    listed = dict(reversed(line.split()) for line in (FIXTURES / "SHA1SUMS").read_text().splitlines())
    assert set(listed) == {path.name for path in FIXTURES.iterdir()} - {"SHA1SUMS", "README.md"}
    for name, digest in listed.items():
        assert hashlib.sha1((FIXTURES / name).read_bytes()).hexdigest() == digest, name


@pytest.mark.parametrize("name", REPAIRED)
def test_each_part_says_so_once_on_its_warnings(name: str) -> None:
    package = load(source(name))
    every_part_read(package)
    records = [w for w in package.warnings if isinstance(w, UndeclaredPrefix)]
    assert {record.part_name: record.prefixes for record in records} == REPAIRED[name]
    record = records[0]
    assert record.code == "mce.undeclared_prefix" and record.detail == " ".join(record.prefixes)
    assert record.to_dict() == {"code": record.code, "part_name": record.part_name, "prefixes": list(record.prefixes), "message": record.message}


@pytest.mark.parametrize("name", REPAIRED)
def test_the_text_model_raises_it_once_a_part_and_reads_what_word_read(name: str) -> None:
    result = extract(load(source(name)))
    raised = [(w.selector, w.detail) for w in result.warnings if w.kind == TextWarning.MCE_UNDECLARED_PREFIX]
    assert sorted(raised) == sorted((part, " ".join(prefixes)) for part, prefixes in REPAIRED[name].items())
    assert all(w.address is None for w in result.warnings if w.kind == "mce.undeclared_prefix")
    # Word's reading is its control's, and so is this package's
    assert word_text(name) == word_text("ui-00-control")
    assert plain_text(result) == plain_text(extract(load(source("ui-00-control"))))


@pytest.mark.parametrize("name", SILENT)
def test_a_declared_prefix_says_nothing(name: str) -> None:
    package = load(source(name))
    assert not [w for w in extract(package).warnings if w.kind == "mce.undeclared_prefix"]
    every_part_read(package)
    assert not [w for w in package.warnings if isinstance(w, UndeclaredPrefix)]


def test_a_numbering_part_the_emulator_reads_with_lxml_is_asked_too() -> None:
    """The property reads the root's start tag from the bytes, for a reader that never unmarshals the part."""
    package = load(source("ui-12-styles"))
    part = package.style_definitions_part
    assert part.undeclared_ignorable == ("zz9",) and part._contents is None


@pytest.mark.parametrize("name", REPAIRED)
def test_a_save_writing_each_part_anew_drops_the_prefix_and_reloads_with_nothing_said(name: str) -> None:
    package = load(source(name))
    every_part_read(package)
    data = saved(package)
    assert undeclared(data) == {}
    again = load(data)
    every_part_read(again)
    assert not [w for w in again.warnings if isinstance(w, UndeclaredPrefix)]
    assert plain_text(extract(again)) == plain_text(extract(load(source("ui-00-control"))))


def test_the_rest_of_the_list_stays() -> None:
    xml = zipfile.ZipFile(io.BytesIO(source("ui-01-main-unknown"))).read("word/document.xml")
    assert etree.fromstring(xml).get(IGNORABLE) == "zz9"
    package = load(source("ui-01-main-unknown"))
    root = package.main_document_part
    assert "w14" in etree.fromstring(xml).nsmap  # the control's root declares it
    root.set_bytes(xml.replace(b'mc:Ignorable="zz9"', b'mc:Ignorable="w14 zz9"'))
    root.contents
    written = etree.fromstring(zipfile.ZipFile(io.BytesIO(saved(package))).read("word/document.xml"))
    assert written.get(IGNORABLE) == "w14"


def test_a_list_left_empty_goes() -> None:
    package = load(source("ui-01-main-unknown"))
    package.main_document_part.contents
    written = etree.fromstring(zipfile.ZipFile(io.BytesIO(saved(package))).read("word/document.xml"))
    assert written.get(IGNORABLE) is None


def test_a_prefix_the_part_now_declares_stays() -> None:
    """``ui-04``: a paragraph declares ``w14`` for its own ``w14:paraId``, so the saved root declares it, and lists it."""
    package = load(source("ui-04-main-lower-known"))
    package.main_document_part.contents
    written = etree.fromstring(zipfile.ZipFile(io.BytesIO(saved(package))).read("word/document.xml"))
    assert written.get(IGNORABLE) == "w14" and "w14" in written.nsmap
    assert b'w14:paraId="1A2B3C4D"' in etree.tostring(written)


def test_a_known_prefix_is_dropped_too_not_declared_from_the_table() -> None:
    """``ui-02``: nothing is in ``w14``, so the list goes without it, where the save declared it before."""
    package = load(source("ui-02-main-known"))
    package.main_document_part.contents
    written = etree.fromstring(zipfile.ZipFile(io.BytesIO(saved(package))).read("word/document.xml"))
    assert written.get(IGNORABLE) is None and "w14" not in written.nsmap


@pytest.mark.parametrize("name", ["ui-01-main-unknown", "ui-12-styles"])
def test_a_save_of_the_source_bytes_keeps_it(name: str) -> None:
    package = load(source(name))
    extract(package)  # a reading is private: nothing is marked for writing
    assert undeclared(saved(package)) == undeclared(source(name)) != {}


def test_a_descendant_declaring_the_prefix_doesn_t_put_it_in_scope() -> None:
    """``ui-05``, which Word refuses: the element declaring ``zz9`` for itself isn't ignorable. Read, as REQ-208 allows,
    the element listed in ``part.skipped``, and the warning raised."""
    package = load(source("ui-05-main-lower-unknown"))
    result = extract(package)
    assert [(w.selector, w.detail) for w in result.warnings if w.kind == "mce.undeclared_prefix"] == [("/word/document.xml", "zz9")]
    package.main_document_part.contents
    assert [etree.QName(item.qname).localname for item in package.main_document_part.skipped] == ["x"]
    assert "UI-ZZ9" not in plain_text(result)


def test_set_bytes_asks_again() -> None:
    package = load(source("ui-01-main-unknown"))
    part = package.main_document_part
    assert part.undeclared_ignorable == ("zz9",)
    part.set_bytes(zipfile.ZipFile(io.BytesIO(source("ui-00-control"))).read("word/document.xml"))
    assert part.undeclared_ignorable == ()
