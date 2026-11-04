"""A package holding two entries of one name is refused; ``duplicate_entries="first"`` reads the first. CR-032 §5.

WD REQ-206: two names are one where they are equivalent as OPC part names,
compared without case, and the rule covers every entry, ``[Content_Types].xml``
and the relationships parts included. Before CR-032 the store read the last
copy silently, and an untouched save dropped the first (section 1).

Every zip is built here, from a package this package makes (one paragraph, a
header and an image), the first copy of a name holding ``FIRST`` and the second
``SECOND``. Over each case:

- the default load refuses it, :class:`DuplicateEntryException` (an
  :class:`InvalidFormatException`) naming every duplicated name in listing
  order with its count, before anything is read (D4);
- ``LoadOptions(duplicate_entries="first")`` reads the first entry of each name
  in the central directory, the relationship's spelling or not (Word's check
  75a), records one :class:`DuplicateEntry` per name, first on
  ``package.warnings``, and logs it (D6); ``strict=True`` does not refuse it;
- the save under it writes only the copy read, with no change to ``save.py``
  (D7): an untouched part's bytes, the content types and the relationships are
  the first's, and an entry no relationship names is not carried.

Then the key (D1, check 75's pairs, section 3.1), every store (D2), the zip
store's first ``ZipInfo`` (D3), the option's refusal (D5), ``insert_ooxml``'s
flat OPC (D9), and every sample's save, which holds no duplicate.
"""

from __future__ import annotations

import functools
import io
import logging
import warnings
import zipfile
from pathlib import Path
from typing import Any

import pytest

from docx4j_py import create_package, load
from docx4j_py.fragments import wml
from docx4j_py.openpackaging import DuplicateEntry, DuplicateEntryException
from docx4j_py.openpackaging.breaks import BrOutsideRun
from docx4j_py.openpackaging.exceptions import Docx4JException, InvalidFormatException
from docx4j_py.openpackaging.load import LoadOptions
from docx4j_py.openpackaging.part_name import PartName
from docx4j_py.openpackaging.parts.binary_part import ImagePart
from docx4j_py.openpackaging.parts.wml import HeaderPart
from docx4j_py.openpackaging.stores import (
    DirectoryPartStore,
    FlatOpcStore,
    MemoryPartStore,
    PartStore,
    ZipPartStore,
)
from docx4j_py.traversal import text_of

ROOT = Path(__file__).resolve().parents[2]
SAMPLES = ROOT / "samples"
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
DOC = "word/document.xml"
HEADER = "word/header1.xml"
IMAGE = "word/media/image1.png"
CONTENT_TYPES = "[Content_Types].xml"
PACKAGE_RELS = "_rels/.rels"
DOCUMENT_RELS = "word/_rels/document.xml.rels"
MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
TEMPLATE = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
FIRST = LoadOptions(duplicate_entries="first")


# ---------------------------------------------------------------------------
# the zips
# ---------------------------------------------------------------------------


def _saved(pkg: Any) -> bytes:
    out = io.BytesIO()
    pkg.save(out)
    return out.getvalue()


def _entries(data: bytes) -> list[tuple[str, bytes]]:
    """Every entry of a zip, duplicates included, in central-directory order."""
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return [(info.filename, z.read(info)) for info in z.infolist()]


def _zip(entries: list[tuple[str, bytes]]) -> bytes:
    """A zip of these entries in this order, a name repeated as given (``zipfile`` warns, and writes it)."""
    out = io.BytesIO()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in entries:
                z.writestr(name, data)
    return out.getvalue()


@functools.cache
def _base() -> tuple[tuple[str, bytes], ...]:
    """The package every case starts from, as entries: a paragraph, MARKER; a header, HEADER MARKER; an image."""
    pkg = create_package()
    mdp = pkg.main_document_part
    body = mdp.contents.body
    body.content.append(wml(f"<w:p {W}><w:r><w:t>MARKER</w:t></w:r></w:p>"))
    header = HeaderPart()
    rel = mdp.add_target_part(header)
    header.set_contents(wml(f"<w:hdr {W}><w:p><w:r><w:t>HEADER MARKER</w:t></w:r></w:p></w:hdr>"))
    body.sect_pr.header_reference_or_footer_reference.append(
        wml(f'<w:headerReference {W} {R} w:type="default" r:id="{rel.id}"/>')
    )
    image = ImagePart(f"/{IMAGE}")
    image.set_bytes(b"\x89PNG MARKER")
    mdp.add_target_part(image)
    return tuple(_entries(_saved(pkg)))


def _data(name: str) -> bytes:
    return dict(_base())[name]


def _marked(name: str, word: bytes) -> bytes:
    return _data(name).replace(b"MARKER", word)


def _swap(name: str, *entries: tuple[str, bytes], extra: tuple[tuple[str, bytes], ...] = ()) -> bytes:
    """The base as a zip, `name`'s entry replaced by `entries` in place, and `extra` at the end."""
    out: list[tuple[str, bytes]] = []
    for entry, data in _base():
        out.extend(entries if entry == name else ((entry, data),))
    return _zip([*out, *extra])


def _doubled(name: str, *spellings: str) -> bytes:
    """`name` FIRST, then a copy for each spelling given (`name` itself by default): SECOND, THIRD."""
    words = (b"SECOND", b"THIRD")
    copies = ((s, _marked(name, w)) for s, w in zip(spellings or (name,), words))
    return _swap(name, (name, _marked(name, b"FIRST")), *copies)


def _text(pkg: Any) -> str:
    return text_of(pkg.main_document_part.contents.body)


def _header_text(pkg: Any) -> str:
    (header,) = [p for p in pkg.parts.parts() if isinstance(p, HeaderPart)]
    return text_of(header.contents)


def _names(data: bytes, name: str) -> list[str]:
    """The saved entries whose name is `name`, compared without case."""
    return [entry for entry, _ in _entries(data) if entry.lower() == name.lower()]


def _saved_entry(data: bytes, name: str) -> bytes:
    (found,) = [d for entry, d in _entries(data) if entry.lower() == name.lower()]
    return found


# ---------------------------------------------------------------------------
# the main part: the same name twice, a case variant, three copies
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("spellings", "named", "count"),
    [
        ((), DOC, 2),
        (("Word/Document.xml",), DOC, 2),
        ((DOC, "WORD/document.XML"), DOC, 3),
    ],
    ids=["the same name twice", "a case variant", "three copies"],
)
def test_the_main_part_twice_is_refused_naming_it(spellings, named, count):
    with pytest.raises(DuplicateEntryException) as raised:
        load(_doubled(DOC, *spellings))
    error = raised.value
    assert isinstance(error, InvalidFormatException) and isinstance(error, Docx4JException)
    assert error.code == "package.duplicate_entry"
    assert error.entries == {named: count}
    assert str(error) == (
        f"Duplicate entry {named} in this package: {count} entries (REQ-206). "
        'LoadOptions(duplicate_entries="first") reads the first'
    )


@pytest.mark.parametrize(
    ("spellings", "count"),
    [((), 2), (("Word/Document.xml",), 2), ((DOC, "WORD/document.XML"), 3)],
    ids=["the same name twice", "a case variant", "three copies"],
)
def test_under_first_the_main_parts_first_copy_is_read_and_saved(spellings, count):
    pkg = load(_doubled(DOC, *spellings), options=FIRST)
    assert _text(pkg) == "FIRST"
    assert pkg.warnings == [DuplicateEntry(DOC, count)]
    assert pkg.skipped == []
    saved = _saved(pkg)
    assert _names(saved, DOC) == [DOC]
    assert b"FIRST" in _saved_entry(saved, DOC) and b"SECOND" not in _saved_entry(saved, DOC)


def test_an_untouched_save_under_first_writes_the_first_copys_bytes():
    pkg = load(_doubled(DOC, "Word/Document.xml"), options=FIRST)
    saved = _saved(pkg)
    assert not pkg.main_document_part.is_unmarshalled
    assert _saved_entry(saved, DOC) == _marked(DOC, b"FIRST")
    assert len(_entries(saved)) == len(_base())


def test_the_decoy_first_is_the_copy_read():
    """``Word/Document.xml`` first, the relationship naming the second's spelling exactly: the first is read (75a)."""
    data = _swap(DOC, ("Word/Document.xml", _marked(DOC, b"FIRST")), (DOC, _marked(DOC, b"SECOND")))
    with pytest.raises(DuplicateEntryException) as raised:
        load(data)
    assert raised.value.entries == {"Word/Document.xml": 2}
    pkg = load(data, options=FIRST)
    assert _text(pkg) == "FIRST"
    assert pkg.warnings == [DuplicateEntry("Word/Document.xml", 2)]
    saved = _saved(pkg)
    assert _names(saved, DOC) == [DOC] and b"FIRST" in _saved_entry(saved, DOC)


# ---------------------------------------------------------------------------
# check 75a: word/headerA.xml then word/headera.xml, the relationship naming headera.xml
# ---------------------------------------------------------------------------


def _check_75a() -> bytes:
    out: list[tuple[str, bytes]] = []
    for name, data in _base():
        if name == HEADER:
            out.append(("word/headerA.xml", data.replace(b"HEADER MARKER", b"HEADER A")))
            out.append(("word/headera.xml", data.replace(b"HEADER MARKER", b"HEADER a")))
        else:
            out.append((name, data.replace(b"header1.xml", b"headera.xml")))  # the relationship and the override
    return _zip(out)


def test_check_75a_is_refused():
    with pytest.raises(DuplicateEntryException) as raised:
        load(_check_75a())
    assert raised.value.entries == {"word/headerA.xml": 2}


def test_check_75a_under_first_reads_the_first_entrys_header_as_word_does():
    pkg = load(_check_75a(), options=FIRST)
    assert _header_text(pkg) == "HEADER A"  # Word showed headerA's header, the relationship naming headera.xml
    assert pkg.warnings == [DuplicateEntry("word/headerA.xml", 2)]
    saved = _saved(pkg)
    assert _names(saved, "word/headera.xml") == ["word/headera.xml"]
    assert b"HEADER A" in _saved_entry(saved, "word/headera.xml")


# ---------------------------------------------------------------------------
# the content types and the relationships parts
# ---------------------------------------------------------------------------


def test_the_content_types_twice_the_firsts_hold():
    second = _data(CONTENT_TYPES).replace(MAIN.encode(), TEMPLATE.encode())
    assert second != _data(CONTENT_TYPES)
    data = _swap(CONTENT_TYPES, (CONTENT_TYPES, _data(CONTENT_TYPES)), (CONTENT_TYPES, second))
    with pytest.raises(DuplicateEntryException) as raised:
        load(data)
    assert raised.value.entries == {CONTENT_TYPES: 2}
    pkg = load(data, options=FIRST)
    assert pkg.content_type_manager.get_content_type(PartName(f"/{DOC}")) == MAIN
    assert pkg.main_document_part.content_type == MAIN
    saved = _saved(pkg)
    assert _names(saved, CONTENT_TYPES) == [CONTENT_TYPES]
    assert MAIN.encode() in _saved_entry(saved, CONTENT_TYPES)
    assert TEMPLATE.encode() not in _saved_entry(saved, CONTENT_TYPES)


def test_the_package_relationships_twice_the_firsts_main_part():
    second = _data(PACKAGE_RELS).replace(b'Target="word/document.xml"', b'Target="word/other.xml"')
    assert second != _data(PACKAGE_RELS)
    other = ("word/other.xml", _marked(DOC, b"SECOND"))
    data = _swap(PACKAGE_RELS, (PACKAGE_RELS, _data(PACKAGE_RELS)), (PACKAGE_RELS, second), extra=(other,))
    with pytest.raises(DuplicateEntryException) as raised:
        load(data)
    assert raised.value.entries == {PACKAGE_RELS: 2}
    pkg = load(data, options=FIRST)
    assert pkg.main_document_part.part_name == PartName(f"/{DOC}")
    assert _text(pkg) == "MARKER"
    saved = _saved(pkg)
    assert _names(saved, PACKAGE_RELS) == [PACKAGE_RELS]
    assert b"word/document.xml" in _saved_entry(saved, PACKAGE_RELS)
    assert _names(saved, DOC) == [DOC] and _names(saved, "word/other.xml") == []


def test_a_parts_relationships_twice_the_firsts():
    second = _data(DOCUMENT_RELS).replace(b"header1.xml", b"header2.xml")
    assert second != _data(DOCUMENT_RELS)
    header2 = ("word/header2.xml", _data(HEADER).replace(b"HEADER MARKER", b"HEADER SECOND"))
    data = _swap(DOCUMENT_RELS, (DOCUMENT_RELS, _data(DOCUMENT_RELS)), (DOCUMENT_RELS, second), extra=(header2,))
    with pytest.raises(DuplicateEntryException) as raised:
        load(data)
    assert raised.value.entries == {DOCUMENT_RELS: 2}
    pkg = load(data, options=FIRST)
    assert _header_text(pkg) == "HEADER MARKER"
    saved = _saved(pkg)
    assert _names(saved, DOCUMENT_RELS) == [DOCUMENT_RELS]
    assert b"header1.xml" in _saved_entry(saved, DOCUMENT_RELS)
    assert _names(saved, HEADER) == [HEADER] and _names(saved, "word/header2.xml") == []


# ---------------------------------------------------------------------------
# every entry: a media binary, and an entry no relationship names
# ---------------------------------------------------------------------------


def test_a_media_binary_twice_is_refused_and_its_first_copy_saved_under_first():
    data = _doubled(IMAGE)
    with pytest.raises(DuplicateEntryException) as raised:
        load(data)
    assert raised.value.entries == {IMAGE: 2}
    pkg = load(data, options=FIRST)
    assert pkg.get_part(f"/{IMAGE}").data == b"\x89PNG FIRST"
    assert pkg.warnings == [DuplicateEntry(IMAGE, 2)]
    saved = _saved(pkg)
    assert _names(saved, IMAGE) == [IMAGE]
    assert _saved_entry(saved, IMAGE) == b"\x89PNG FIRST"


def test_an_entry_no_relationship_names_twice_is_refused_and_not_carried_under_first():
    data = _zip([*_base(), ("word/orphan.xml", b"<FIRST/>"), ("word/orphan.xml", b"<SECOND/>")])
    with pytest.raises(DuplicateEntryException) as raised:
        load(data)
    assert raised.value.entries == {"word/orphan.xml": 2}
    pkg = load(data, options=FIRST)
    assert pkg.warnings == [DuplicateEntry("word/orphan.xml", 2)]
    assert _names(_saved(pkg), "word/orphan.xml") == []


def test_several_names_are_listed_in_one_refusal_in_listing_order():
    rels = (PACKAGE_RELS, _data(PACKAGE_RELS))
    data = _swap(DOC, (DOC, _marked(DOC, b"FIRST")), (DOC, _marked(DOC, b"SECOND")), extra=(rels, rels))
    with pytest.raises(DuplicateEntryException) as raised:
        load(data)
    assert raised.value.entries == {PACKAGE_RELS: 3, DOC: 2}  # _rels/.rels is listed first
    assert str(raised.value) == (
        f"Duplicate entries in this package: {PACKAGE_RELS}, 3 entries; {DOC}, 2 entries (REQ-206). "
        'LoadOptions(duplicate_entries="first") reads the first of each'
    )
    pkg = load(data, options=FIRST)
    assert pkg.warnings == [DuplicateEntry(PACKAGE_RELS, 3), DuplicateEntry(DOC, 2)]


# ---------------------------------------------------------------------------
# the refusal reads nothing; the warning; strict; the option
# ---------------------------------------------------------------------------


class _CountingStore(ZipPartStore):
    """A zip store that records what is loaded from it."""

    def __init__(self, data: bytes) -> None:
        super().__init__(data)
        self.loaded: list[str] = []

    def load(self, part_name: str) -> bytes:
        self.loaded.append(part_name)
        return super().load(part_name)


def test_the_refusal_comes_before_anything_is_read():
    store = _CountingStore(_doubled(DOC))
    with pytest.raises(DuplicateEntryException):
        load(store)
    assert store.loaded == []


def test_the_warning_is_first_on_package_warnings_and_logged(caplog):
    first = _data(DOC).replace(b"<w:r><w:t>MARKER</w:t></w:r>", b"<w:br/><w:r><w:t>FIRST</w:t></w:r>")
    assert first != _data(DOC)
    data = _swap(DOC, (DOC, first), (DOC, _marked(DOC, b"SECOND")))
    with caplog.at_level(logging.WARNING, logger="docx4j_py.openpackaging"):
        pkg = load(data, options=FIRST)
    record = DuplicateEntry(DOC, 2)
    assert record.message in caplog.messages
    assert pkg.warnings == [record]
    assert _text(pkg) == "\nFIRST"  # the parse read the first copy's w:br outside a w:r (CR-030)
    found = pkg.warnings
    assert found[0] == record and isinstance(found[1], BrOutsideRun) and len(found) == 2
    assert record.code == DuplicateEntry.code == DuplicateEntryException.code == "package.duplicate_entry"
    assert record.to_dict() == {
        "code": "package.duplicate_entry",
        "entry": DOC,
        "count": 2,
        "message": f"Duplicate entry {DOC} in this package: 2 entries (REQ-206); the first read",
    }
    assert str(record) == record.message


def test_strict_does_not_refuse_under_first():
    pkg = load(_doubled(DOC), options=LoadOptions(strict=True, duplicate_entries="first"))
    assert _text(pkg) == "FIRST"
    assert pkg.warnings == [DuplicateEntry(DOC, 2)]


def test_strict_refuses_by_default():
    with pytest.raises(DuplicateEntryException):
        load(_doubled(DOC), options=LoadOptions(strict=True))


@pytest.mark.parametrize("value", ["last", "FIRST", "", None])
def test_another_value_is_refused_naming_req_206(value):
    with pytest.raises(ValueError, match=r"duplicate_entries must be one of \('refuse', 'first'\).*REQ-206"):
        LoadOptions(duplicate_entries=value)


def test_the_default_is_refuse():
    assert LoadOptions().duplicate_entries == "refuse"
    assert create_package().warnings == []


def test_a_clean_package_has_no_duplicates_and_no_warning():
    data = _zip(list(_base()))
    assert ZipPartStore(data).duplicates() == {}
    assert load(data).warnings == []
    assert load(data, options=FIRST).warnings == []


# ---------------------------------------------------------------------------
# the key (D1): check 75's pairs, backslashes, percent-encoding
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("first", "second", "duplicates"),
    [
        ("word/headerA.xml", "word/headera.xml", {"word/headerA.xml": 2}),  # check 75a
        ("word/header%C3%89.xml", "word/header%C3%A9.xml", {}),  # two names: the hex digits' letters fold, not É
        ("word/header%c3%a9.xml", "word/header%C3%A9.xml", {"word/header%c3%a9.xml": 2}),  # one
        ("word/headerÉ.xml", "word/headeré.xml", {"word/headerÉ.xml": 2}),  # one: str.lower() folds É (D1's departure)
        ("word/header%C3%A9.xml", "word/headeré.xml", {}),  # encoded and decoded: two names, never confused
        ("word/document.xml", "word\\document.xml", {"word/document.xml": 2}),  # a backslash folded
        ("/word/document.xml", "word/document.xml", {"word/document.xml": 2}),  # the leading slash dropped
    ],
    ids=["75a", "%C3%89 and %C3%A9", "%c3%a9 and %C3%A9", "raw É and é", "encoded and decoded", "backslash", "slash"],
)
def test_the_key_is_the_stores_own(first, second, duplicates):
    store = ZipPartStore(_zip([(first, b"FIRST"), (second, b"SECOND")]))
    assert store.duplicates() == duplicates
    assert len(store.part_names()) == (1 if duplicates else 2)
    assert store.load(first) == b"FIRST"
    if duplicates:
        assert store.load(second) == b"FIRST"  # one name: the first entry, by either spelling
    else:
        assert store.load(second) == b"SECOND"


def test_a_backslash_twin_of_the_main_part_is_refused():
    with pytest.raises(DuplicateEntryException) as raised:
        load(_doubled(DOC, "word\\document.xml"))
    assert raised.value.entries == {DOC: 2}


def test_an_encoded_and_a_decoded_spelling_are_not_refused():
    data = _zip([*_base(), ("word/media/%C3%A9.png", b"ONE"), ("word/media/é.png", b"TWO")])
    assert load(data).warnings == []


# ---------------------------------------------------------------------------
# the zip store (D3): the first ZipInfo, by which it reads and sizes; reopen()
# ---------------------------------------------------------------------------


def test_the_zip_store_reads_and_sizes_the_first_entry_and_lists_it_once():
    data = _zip([(DOC, b"FIRST"), ("Word/Document.xml", b"SECOND!"), (DOC, b"THIRD!!!")])
    store = ZipPartStore(data)
    assert store.part_names() == [DOC]
    assert store.has("WORD/DOCUMENT.XML")
    assert store.load(DOC) == store.load("Word/Document.xml") == b"FIRST"
    assert store.size(DOC) == len(b"FIRST")
    assert store.duplicates() == {DOC: 3}
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        assert z.read(DOC) == b"THIRD!!!"  # zipfile's lookup by name finds the last


def test_reopen_rebuilds_the_index_the_same_way(tmp_path):
    path = tmp_path / "dup.docx"
    path.write_bytes(_zip([(DOC, b"FIRST"), (DOC, b"SECOND")]))
    store = ZipPartStore(path)
    try:
        store.reopen()
        assert store.duplicates() == {DOC: 2}
        assert store.load(DOC) == b"FIRST" and store.size(DOC) == 5
    finally:
        store.close()


# ---------------------------------------------------------------------------
# the other stores (D2)
# ---------------------------------------------------------------------------


def test_a_directory_with_case_variant_files(tmp_path):
    for name, data in _base():
        target = tmp_path / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data.replace(b"MARKER", b"SECOND") if name == DOC else data)
    (tmp_path / "word" / "Document.xml").write_bytes(_marked(DOC, b"FIRST"))
    if len(list((tmp_path / "word").glob("*ocument.xml"))) != 2:
        pytest.skip("this file system folds case")
    store = DirectoryPartStore(tmp_path)
    # the walk's order is sorted: Document.xml before document.xml
    assert store.duplicates() == {"word/Document.xml": 2}
    with pytest.raises(DuplicateEntryException) as raised:
        load(DirectoryPartStore(tmp_path))
    assert raised.value.entries == {"word/Document.xml": 2}
    pkg = load(DirectoryPartStore(tmp_path), options=FIRST)
    assert _text(pkg) == "FIRST"
    assert pkg.warnings == [DuplicateEntry("word/Document.xml", 2)]


def test_a_memory_store_counts_its_constructors_mapping_and_a_put_replaces():
    store = MemoryPartStore({DOC: b"FIRST", "Word/Document.xml": b"SECOND"})
    assert store.duplicates() == {DOC: 2}
    assert store.part_names() == [DOC] and store.load("WORD/document.xml") == b"FIRST"
    store.put("Word/Document.xml", b"PUT")  # a sink's write: replaced, its first spelling kept, nothing counted
    assert store.part_names() == [DOC] and store.load(DOC) == b"PUT"
    assert store.duplicates() == {DOC: 2}
    fresh = MemoryPartStore()
    fresh.put(DOC, b"ONE")
    fresh.put("Word/Document.xml", b"TWO")
    assert fresh.duplicates() == {} and fresh.part_names() == [DOC] and fresh.load(DOC) == b"TWO"


def test_a_memory_store_with_a_case_variant_is_refused():
    parts = dict(_base())
    parts["Word/Document.xml"] = _marked(DOC, b"SECOND")
    parts[DOC] = _marked(DOC, b"FIRST")
    with pytest.raises(DuplicateEntryException) as raised:
        load(MemoryPartStore(parts))
    assert raised.value.entries == {DOC: 2}
    assert _text(load(MemoryPartStore(parts), options=FIRST)) == "FIRST"


def _flat(*parts: tuple[str, str, str]) -> str:
    """A ``pkg:package`` of (name, content type, XML) parts, in order."""
    body = "".join(
        f'<pkg:part pkg:name="{name}" pkg:contentType="{content_type}"><pkg:xmlData>{xml}</pkg:xmlData></pkg:part>'
        for name, content_type, xml in parts
    )
    return f'<pkg:package xmlns:pkg="{PKG}">{body}</pkg:package>'


def _flat_doubled() -> str:
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Target="word/document.xml" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
        "</Relationships>"
    )
    document = f"<w:document {W}><w:body><w:p><w:r><w:t>%s</w:t></w:r></w:p></w:body></w:document>"
    return _flat(
        ("/_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml", rels),
        ("/word/document.xml", MAIN, document % "FIRST"),
        ("/word/document.xml", TEMPLATE, document % "SECOND"),
    )


def test_a_flat_opc_package_with_two_parts_of_one_name():
    store = FlatOpcStore.parse(_flat_doubled())
    assert store.duplicates() == {DOC: 2}
    with pytest.raises(DuplicateEntryException) as raised:
        load(FlatOpcStore.parse(_flat_doubled()))
    assert raised.value.entries == {DOC: 2}
    pkg = load(FlatOpcStore.parse(_flat_doubled()), options=FIRST)
    assert _text(pkg) == "FIRST"
    assert pkg.main_document_part.content_type == MAIN  # the first's content type, the second's not taken
    assert pkg.warnings == [DuplicateEntry(DOC, 2)]


def test_insert_ooxml_refuses_a_flat_opc_fragment_with_two_parts_of_one_name():
    """CR-003 section 12's one code for a package it cannot read, the refusal quoted (D9)."""
    from docx4j_py.model.content import ContentError

    pkg = create_package()
    with pytest.raises(ContentError) as raised:
        pkg.body.insert_ooxml(_flat_doubled())
    assert raised.value.code == "ooxml.no_main_part"
    assert f"Duplicate entry {DOC} in this package: 2 entries (REQ-206)" in raised.value.message
    assert isinstance(raised.value.__cause__, DuplicateEntryException)
    assert text_of(pkg.main_document_part.contents.body) == ""


class _ThirdPartyStore:
    """A store written elsewhere: the protocol's four methods, no ``duplicates()``."""

    def __init__(self, data: bytes) -> None:
        self._inner = ZipPartStore(data)

    def part_names(self) -> list[str]:
        return self._inner.part_names()

    def has(self, part_name: str) -> bool:
        return self._inner.has(part_name)

    def load(self, part_name: str) -> bytes:
        return self._inner.load(part_name)

    def size(self, part_name: str) -> int | None:
        return self._inner.size(part_name)


def test_a_store_without_duplicates_is_not_checked():
    store = _ThirdPartyStore(_doubled(DOC))
    assert isinstance(store, PartStore)  # duplicates() is not a member of the protocol
    pkg = load(store)
    assert _text(pkg) == "FIRST" and pkg.warnings == []


# ---------------------------------------------------------------------------
# our own saves write no duplicate
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("path", sorted(p for p in SAMPLES.iterdir() if p.suffix == ".docx"), ids=lambda p: p.name)
def test_every_sample_saved_and_re_read_records_no_duplicate(path):
    assert ZipPartStore(path).duplicates() == {}
    with load(path) as pkg:
        saved = pkg.save()
    assert ZipPartStore(saved).duplicates() == {}
    with load(saved) as again:
        assert not [w for w in again.warnings if isinstance(w, DuplicateEntry)]
