"""An OLE compound file's one stream, and an OLE 1.0 package's names: ``model/text/ole.py``. CR-027 sections 4.3 and 5.

The reader's oracle is POI's reading (docx4j's copy, ``Ole10Native`` and
``POIFSFileSystem``), through ``EmbeddedObjectsTest``: ``probe41.bin``, which POI
wrote, and each rule below, which ``Ole10Native(byte[], int)`` holds. The bytes
are ``ole_writer``'s, as section 5 recommends: no test here needs a Java run. Two
packages Word wrote are read from the sibling checkout, where it is there.
"""

from __future__ import annotations

import random
import struct
import zipfile
from pathlib import Path

import pytest
from ole_writer import OLE_MARKER, compound_file, ole10, ole10_record

from docx4j_py.model.text.ole import OLE10_NATIVE, package_names, record_names, root_stream

ROOT = Path(__file__).resolve().parent.parent.parent
PROBE = ROOT / "tests" / "fixtures" / "text-extraction" / "phase0" / "extraction-probe.docx"

DOCX4J_OLE = ROOT.parent / "docx4j" / "docx4j-core-tests" / "src" / "test" / "resources" / "OLE"

NAMES = ("notes.txt", "C:\\docs\\notes.txt", "C:\\tmp\\notes.txt")


def _probe41() -> bytes:
    with zipfile.ZipFile(PROBE) as z:
        return z.read("word/embeddings/probe41.bin")


def _in_a_file(record: bytes, **more) -> bytes:
    return compound_file({"\x01Ole": OLE_MARKER, OLE10_NATIVE: record}, **more)


def record_names_or_none(record: bytes):
    """As :func:`package_names` reads a record: a failure is none."""
    return package_names(compound_file({OLE10_NATIVE: record}))


# --- the compound file ---------------------------------------------------------------------------------


def test_the_file_poi_wrote():
    """``probe41.bin``: POI's own writing, the mini stream's chains and directory as it lays them out."""
    data = _probe41()
    assert package_names(data) == ("probe.txt", "probe.txt", "probe.txt")
    assert root_stream(data, "\x01Ole") == OLE_MARKER
    assert b"OLE41_NATIVE_TEXT" in root_stream(data, OLE10_NATIVE)


def _word_wrote(name: str) -> bytes:
    with zipfile.ZipFile(DOCX4J_OLE / f"{name}.docx") as z:
        return z.read("word/embeddings/oleObject1.bin")


@pytest.mark.skipif(not DOCX4J_OLE.exists(), reason="needs the sibling checkout ../docx4j")
def test_the_package_word_wrote():
    """docx4j's ``OLE/dragged txt.docx``: a text file dragged into Word, its ``CompObj``, ``ObjInfo`` and UTF-16 names.

    The expectation is POI's reading of the same bytes, measured through
    docx4j's copy: ``parsed``, each name as here, ``label2`` and the others passed over.
    """
    assert package_names(_word_wrote("dragged txt")) == (
        "simple.txt",
        "X:\\ole_tests\\simple.txt",
        (
            "C:\\Users\\jharrop\\AppData\\Local\\Temp\\{0AC6109D-2421-4E57-8F7E-933AD9AFB804}"
            "\\{BD445197-EDA4-4464-8E1E-5BFF2E14DB1C}\\simple.txt"
        ),
    )


@pytest.mark.skipif(not DOCX4J_OLE.exists(), reason="needs the sibling checkout ../docx4j")
def test_a_bitmap_word_wrote_has_no_names():
    """docx4j's ``OLE/create new bitmap.docx``: an ``Ole10Native`` stream that is no package's, POI's ``unparsed``."""
    data = _word_wrote("create new bitmap")
    assert len(root_stream(data, OLE10_NATIVE)) == 167172
    assert package_names(data) is None


@pytest.mark.parametrize("shift", [9, 12])
@pytest.mark.parametrize("size", [0, 1, 60, 3000, 4095 - 60, 4096, 9000, 70000])
def test_a_stream_of_any_size(shift, size):
    """Under 4,096 bytes in the mini stream, else in sectors of the file's own size: 512 bytes in version 3, 4,096 in 4."""
    content = bytes(range(256)) * (size // 256 + 1)
    record = ole10_record(*NAMES, content[:size])
    data = _in_a_file(record, shift=shift)
    assert root_stream(data, OLE10_NATIVE) == record
    assert package_names(data) == NAMES


def test_a_fat_past_the_headers_109_sectors():
    """A file big enough for the DIFAT: 109 sectors of 128 entries is 6.8 MB."""
    record = ole10_record(*NAMES, bytes(7_500_000))
    data = _in_a_file(record)
    assert struct.unpack_from("<i", data, 0x48)[0] > 0  # the DIFAT's count: the file does reach it
    assert package_names(data) == NAMES


@pytest.mark.parametrize("name", ["\x01OLE10NATIVE", "\x01ole10native", "\x01Ole10Native"])
def test_the_stream_is_matched_without_case(name):
    """POI's ``getEntryCaseInsensitive``."""
    assert package_names(compound_file({name: ole10_record(*NAMES)})) == NAMES


@pytest.mark.parametrize(
    "data",
    [
        b"",
        b"not an OLE file",
        b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(100),
        compound_file({}),
        compound_file({"\x01Ole": OLE_MARKER}),
        compound_file({"Ole10Native": ole10_record(*NAMES)}),
    ],
    ids=["empty", "text", "cut", "no streams", "marker only", "name without its control"],
)
def test_no_stream_gives_no_names(data):
    assert package_names(data) is None


def test_a_storage_of_that_name_gives_no_names():
    """``createFromEmbeddedOleObject`` casts it to a document and fails."""
    assert package_names(compound_file({"Other": b"x"}, storages=(OLE10_NATIVE,))) is None


def test_a_stream_past_pois_limit_gives_no_names():
    """POI's ``MAX_RECORD_LENGTH``: a size over 100,000,000 is refused, read or not."""
    data = bytearray(_in_a_file(ole10_record(*NAMES)))
    at = _directory_entry(data, OLE10_NATIVE)
    struct.pack_into("<I", data, at + 0x78, 100_000_001)
    assert package_names(bytes(data)) is None


def _directory_entry(data: bytes, name: str) -> int:
    return data.index(name.encode("utf-16-le"))


def test_a_stream_longer_than_its_chain_gives_no_names():
    data = bytearray(_in_a_file(ole10_record(*NAMES)))
    at = _directory_entry(data, OLE10_NATIVE)
    struct.pack_into("<I", data, at + 0x78, 4000)
    assert package_names(bytes(data)) is None


def test_a_chain_that_meets_itself_ends():
    """Each FAT entry points at itself: refused, not walked for ever."""
    data = bytearray(_in_a_file(ole10_record(*NAMES, b"x" * 5000)))
    (fat,) = struct.unpack_from("<I", data, 0x4C)
    fat_at = (fat + 1) * 512
    for i in range(128):
        struct.pack_into("<I", data, fat_at + i * 4, i)
    assert package_names(bytes(data)) is None


def test_a_directory_that_meets_itself_ends():
    """The siblings of a damaged directory point back at each other: the walk ends, with no stream found."""
    data = bytearray(compound_file({"A": b"a", "B": b"b"}))
    a, b = _directory_entry(data, "A"), _directory_entry(data, "B")
    struct.pack_into("<I", data, a + 0x48, 2)
    struct.pack_into("<I", data, b + 0x48, 1)
    assert package_names(bytes(data)) is None


def test_no_file_raises():
    """Nothing here raises, whatever the bytes (section 7's risk), on damaged forms of POI's file and this writer's."""
    rnd = random.Random(27)
    sources = [_probe41(), _in_a_file(ole10_record(*NAMES)), _in_a_file(ole10_record(*NAMES, b"x" * 9000))]
    for _ in range(1500):
        data = bytearray(rnd.choice(sources))
        for _ in range(rnd.randint(1, 4)):
            at = rnd.choice((rnd.randrange(0x74), rnd.randrange(len(data))))
            data[at] = rnd.choice((0, 1, 2, 5, 0xFE, 0xFF, rnd.randrange(256)))
        if rnd.random() < 0.2:
            del data[rnd.randrange(len(data)) :]
        assert package_names(bytes(data)) is None or len(package_names(bytes(data))) == 3


# --- the record ----------------------------------------------------------------------------------------


def test_the_record():
    """``label``, ``fileName`` and ``command``; the data itself is not kept."""
    assert record_names(ole10_record(*NAMES)) == NAMES


def test_an_empty_name_is_absent():
    assert record_names(ole10_record("label", "", "")) == ("label", None, None)


@pytest.mark.parametrize("flags1", [0, 1, 3, -1, 256])
def test_flags1_other_than_2_is_unparsed(flags1):
    assert record_names(ole10_record(*NAMES, flags1=flags1)) is None


@pytest.mark.parametrize("first", [*range(0x01, 0x20), 0x7F])
def test_a_control_first_is_compact(first):
    """``Character.isISOControl``: an equation's record, say, whose first byte is no name's."""
    assert record_names(ole10_record(bytes([first]) + b"abc", "b", "c")) is None


def test_an_empty_label_is_compact():
    """Its first byte is the terminator, 0x00, a control."""
    assert record_names(ole10_record("", "b", "c")) is None


@pytest.mark.parametrize("first", [0x80, 0x9F, 0xC3, 0xFF])
def test_a_high_first_byte_is_no_control(first):
    """Java tests it as a signed byte, so 0x80 to 0x9F, a control as a code point, are negative there."""
    assert record_names(ole10_record(bytes([first]) + b"abc", "b", "c")) is not None


def test_a_name_is_utf8_and_a_command_iso_8859_1():
    """``getFromCompressedUTF8`` and ``getFromCompressedUnicode``."""
    label = "n\u00f6tes.txt".encode()
    assert record_names(ole10_record(label, label, label)) == (
        "n\u00f6tes.txt",
        "n\u00f6tes.txt",
        "n\u00c3\u00b6tes.txt",
    )


@pytest.mark.parametrize(
    ("raw", "text"),
    [
        (b"a\xffb", "a\ufffdb"),
        (b"a\xc3", "a\ufffd"),
        (b"a\xe2\x82", "a\ufffd"),
        (b"a\xed\xa0\x80b", "a\ufffdb"),
        (b"a\xed\xbf\xbf", "a\ufffd"),
        (b"a\xc0\x80", "a\ufffd\ufffd"),
    ],
)
def test_a_malformed_name_decodes_as_java_decodes_it(raw, text):
    """Each measured with JDK 21's ``new String(bytes, UTF_8)``: an encoded surrogate's three bytes are one U+FFFD."""
    assert record_names(ole10_record(raw, "b", "c"))[0] == text


def test_a_name_of_1023_bytes_is_read_and_one_of_1024_is_not():
    """POI's buffer holds the terminator too."""
    assert record_names(ole10_record(b"x" * 1023, "b", "c"))[0] == "x" * 1023
    assert record_names_or_none(ole10_record(b"x" * 1024, "b", "c")) is None
    assert record_names_or_none(ole10_record("a", b"x" * 1024, "c")) is None


def test_a_command_of_1024_bytes_is_read_and_one_of_1025_is_not():
    """The length counts the terminator, and ``toByteArray`` refuses more than 1,024."""
    assert record_names(ole10_record("a", "b", b"x" * 1023))[2] == "x" * 1023
    assert record_names_or_none(ole10_record("a", "b", b"x" * 1024)) is None


def _utf16(text: str) -> bytes:
    return struct.pack("<i", len(text)) + text.encode("utf-16-le")


def test_the_utf16_names_are_read_past_and_not_kept():
    """REQ-512: the plain fields, not ``label2``, ``fileName2`` and ``command2``."""
    after = _utf16("C:\\wide\\cmd") + _utf16("wide.txt") + _utf16("C:\\wide\\wide.txt")
    assert record_names(ole10_record(*NAMES, after=after)) == NAMES


@pytest.mark.parametrize(
    "after",
    [
        b"",
        b"\0",
        struct.pack("<i", 1) + b"ab",
        _utf16("x" * 513) + _utf16("") + _utf16(""),
        _utf16("a") + _utf16("x" * 513) + _utf16(""),
        _utf16("a") + _utf16("b") + _utf16("x" * 513),
        struct.pack("<i", -1) + bytes(40),
        struct.pack("<i", 0x40000001) + bytes(40),
    ],
    ids=[
        "ends at the data",
        "one byte",
        "cut names",
        "first long",
        "second long",
        "third long",
        "negative",
        "overflow",
    ],
)
def test_a_record_that_fails_after_the_data_gives_no_names(after):
    """``ole10Names`` catches the failure: all three absent, read or not. A record ending at its data's end is one."""
    assert record_names_or_none(ole10_record(*NAMES, after=after)) is None


@pytest.mark.parametrize("size", [-1, 0x7FFFFFFF, -0x80000000, 1, 0])
def test_a_total_size_too_short_or_past_an_ints_end_gives_no_names(size):
    assert record_names_or_none(struct.pack("<i", size) + ole10_record(*NAMES)[4:]) is None


def test_a_total_size_past_the_stream_is_the_stream():
    """``limit`` is the lesser: the record is read to the stream's end."""
    assert record_names_or_none(struct.pack("<i", 100000) + ole10_record(*NAMES)[4:]) == NAMES


def test_a_record_too_short_for_its_size_gives_no_names():
    record = ole10_record(*NAMES)
    for end in range(len(record) - 1):
        assert record_names_or_none(record[:end]) is None


def test_ole10_is_docx4js_ole10():
    """The test writer's whole file: docx4j's helper, in the mini stream and in 4,096-byte sectors."""
    assert package_names(ole10(*NAMES, "NOTES_TEXT\r\n")) == NAMES
    assert package_names(ole10(*NAMES, "NOTES_TEXT\r\n", shift=12)) == NAMES
