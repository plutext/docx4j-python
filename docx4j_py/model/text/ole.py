"""An OLE compound file's one stream, and an OLE 1.0 package's names. CR-027 section 4.3 (Q3).

What docx4j's ``Embedded.ole10Names`` reads through its copy of POI
(``org.docx4j.org.apache.poi.poifs``), and no more: the compound file's header,
its FAT and mini FAT, its directory and one stream of the root storage,
``\\x01Ole10Native``, its name matched without case (``getEntryCaseInsensitive``);
then that stream's ``label``, ``fileName`` and ``command``, as
``Ole10Native(byte[], int)`` reads them:

* names only where ``flags1`` is 2 and the label's first byte is not a control
  (0x00 to 0x1F or 0x7F: Java tests it as a signed byte); otherwise the record
  is compact or unparsed and has none;
* ``label`` and ``fileName`` null-terminated within 1,024 bytes and UTF-8, a
  malformed byte U+FFFD;
* ``command`` after its length, ISO-8859-1, its last byte dropped;
* the data, and the UTF-16 ``command2``, ``label2`` and ``fileName2`` after it,
  read past and not kept: a record that fails in them gives no names either,
  and one that ends at its data's end fails there, as POI's ``readShort`` does.

Nothing here raises: a part that is no compound file, holds no such stream or
whose record fails at any point gives None, as ``ole10Names`` catches everything
and returns null. So a host's filter meets ``notes.txt`` where the package has
it and nothing where it does not. Every read is bounded by the part's size: a
chain that meets itself is refused, and a stream is read only as far as its
size, itself limited to POI's 100,000,000 bytes.

No file is opened, and no object's content is read: reading an OLE 1.0 package's
file stays deferred (CR-011 Q7).
"""

from __future__ import annotations

import re
import struct
from collections.abc import Callable, Iterator

__all__ = ["OLE10_NATIVE", "package_names", "record_names", "root_stream"]

#: The stream an OLE 1.0 package's native data is in (POI ``Ole10Native.OLE10_NATIVE``).
OLE10_NATIVE = "\x01Ole10Native"

_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_HEADER = 512
#: A chain's end; and a free sector, which a damaged chain may end at, or a directory entry's "no sibling" or "no child".
_END_OF_CHAIN = 0xFFFFFFFE
_NO_ENTRY = 0xFFFFFFFF
#: A stream shorter than this is in the mini stream, of sectors this long. POI holds both constant, as [MS-CFB] does.
_MINI_CUTOFF = 4096
_MINI_SECTOR = 64
_ENTRY = 128
_STREAM = 2
_HEADER_FAT_SECTORS = 109
#: POI's ``DEFAULT_MAX_RECORD_LENGTH`` and ``DEFAULT_MAX_STRING_LENGTH``.
_MAX_RECORD = 100_000_000
_MAX_STRING = 1024
_INT_MAX = 0x7FFFFFFF
_SURROGATE = re.compile(rb"\xed[\xa0-\xbf][\x80-\xbf]")

Names = tuple[str | None, str | None, str | None]


class _NotRead(Exception):
    """A compound file or a record that gives no names; never leaves this module."""


class _CompoundFile:
    """A compound file, read only as far as one stream of its root storage needs."""

    __slots__ = ("_data", "_directory", "_fat", "_first_mini_fat", "_sector")

    def __init__(self, data: bytes) -> None:
        """The header and the FAT's sectors; raises where it is no compound file this reader can read."""
        if len(data) < _HEADER or data[: len(_SIGNATURE)] != _SIGNATURE:
            raise _NotRead("no compound file")
        shift = data[0x1E]
        if shift not in (9, 12):
            raise _NotRead(f"a sector shift of {shift}")
        self._data = data
        self._sector = 1 << shift
        fat_count, first_directory = struct.unpack_from("<iI", data, 0x2C)
        self._first_mini_fat, _, first_difat, difat_count = struct.unpack_from("<IiIi", data, 0x3C)
        if fat_count < 0:
            raise _NotRead(f"{fat_count} FAT sectors")
        self._fat = self._fat_sectors(fat_count, first_difat, difat_count)
        self._directory = self._chain_bytes(first_directory)

    def _at(self, sector: int) -> bytes:
        """A sector's bytes; the last of a file cut short may be short, as POI's data source reads it."""
        offset = (sector + 1) * self._sector
        if offset >= len(self._data):
            raise _NotRead(f"sector {sector} is past the file's end")
        return self._data[offset : offset + self._sector]

    def _fat_sectors(self, count: int, first_difat: int, difat_count: int) -> list[int]:
        """The sectors holding the FAT: the header's 109, then the DIFAT's, each in the file once."""
        header = struct.unpack_from(f"<{_HEADER_FAT_SECTORS}I", self._data, 0x4C)
        sectors = list(header[: min(count, _HEADER_FAT_SECTORS)])
        remaining = count - len(sectors)
        per_block = self._sector // 4 - 1
        block, seen = first_difat, set()
        for _ in range(difat_count):
            if block in seen:
                raise _NotRead("the DIFAT meets itself")
            seen.add(block)
            *entries, block = struct.unpack(f"<{per_block + 1}I", self._at(block))
            taken = entries[: min(remaining, per_block)]
            remaining -= len(taken)
            sectors.extend(s for s in taken if s not in (_NO_ENTRY, _END_OF_CHAIN))
        if len(set(sectors)) != len(sectors):
            raise _NotRead("a FAT sector listed twice")
        return sectors

    def _next(self, sector: int) -> int:
        """The sector after this one in the FAT."""
        index, entry = divmod(sector, self._sector // 4)
        if index >= len(self._fat):
            raise _NotRead(f"sector {sector} is past the FAT's end")
        return struct.unpack_from("<I", self._at(self._fat[index]), entry * 4)[0]

    @staticmethod
    def _chain(start: int, following: Callable[[int], int]) -> Iterator[int]:
        """A chain's sectors in order; one that meets itself is refused, so a walk is bounded by the file."""
        sector, seen = start, set()
        while sector not in (_END_OF_CHAIN, _NO_ENTRY):
            if sector in seen:
                raise _NotRead("a chain meets itself")
            seen.add(sector)
            yield sector
            sector = following(sector)

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self._at(sector) for sector in self._chain(start, self._next))

    def _entry(self, index: int) -> bytes | None:
        entry = self._directory[index * _ENTRY : (index + 1) * _ENTRY]
        return entry if len(entry) == _ENTRY else None

    def stream(self, name: str) -> bytes | None:
        """The root storage's stream of this name, matched without case; None for none, or an entry that is a storage."""
        root = self._entry(0)
        if root is None:
            raise _NotRead("no root entry")
        wanted = name.upper()
        waiting, seen = [_child(root)], set()
        while waiting:
            index = waiting.pop()
            if index == _NO_ENTRY or index in seen:
                continue
            seen.add(index)
            entry = self._entry(index)
            if entry is None:
                continue
            if _name(entry).upper() == wanted:
                return self._read(entry, root) if entry[0x42] == _STREAM else None
            waiting += struct.unpack_from("<II", entry, 0x44)
        return None

    def _read(self, entry: bytes, root: bytes) -> bytes:
        (start,) = struct.unpack_from("<I", entry, 0x74)
        (size,) = struct.unpack_from("<I", entry, 0x78)
        if size > _MAX_RECORD:
            raise _NotRead(f"a stream of {size} bytes")
        if size < _MINI_CUTOFF:
            return self._read_mini(start, size, root)
        return _gathered((self._at(sector) for sector in self._chain(start, self._next)), size)

    def _read_mini(self, start: int, size: int, root: bytes) -> bytes:
        """A stream in the mini stream, which the root entry's chain holds, by the mini FAT's chain."""
        mini_fat = self._chain_bytes(self._first_mini_fat)
        mini_stream = self._chain_bytes(struct.unpack_from("<I", root, 0x74)[0])

        def following(sector: int) -> int:
            return struct.unpack_from("<I", mini_fat, sector * 4)[0]

        sectors = self._chain(start, following)
        return _gathered((mini_stream[s * _MINI_SECTOR : (s + 1) * _MINI_SECTOR] for s in sectors), size)


def _gathered(pieces: Iterator[bytes], size: int) -> bytes:
    """A stream's `size` bytes, walking its chain no further; a chain that ends sooner is refused."""
    out = bytearray()
    for piece in pieces if size else ():
        out += piece
        if len(out) >= size:
            break
    if len(out) < size:
        raise _NotRead("a stream shorter than its size")
    return bytes(out[:size])


def _name(entry: bytes) -> str:
    """A directory entry's name: its UTF-16 less the terminator its length counts."""
    (length,) = struct.unpack_from("<H", entry, 0x40)
    return entry[: max(0, min(length // 2 - 1, 32)) * 2].decode("utf-16-le", errors="replace")


def _child(entry: bytes) -> int:
    return struct.unpack_from("<I", entry, 0x4C)[0]


class _Record:
    """An ``Ole10Native`` record being read: POI's ``leis``, every read refused past its limit."""

    __slots__ = ("_data", "_end", "at")

    def __init__(self, data: bytes, at: int, end: int) -> None:
        self._data = data
        self.at = at
        self._end = end

    def take(self, size: int) -> bytes:
        if size > self._end - self.at:
            raise _NotRead("past the record's end")
        out = self._data[self.at : self.at + size]
        self.at += size
        return out

    def short(self) -> int:
        return int.from_bytes(self.take(2), "little", signed=True)

    def integer(self) -> int:
        return int.from_bytes(self.take(4), "little", signed=True)

    def sized(self, size: int, limit: int) -> bytes:
        """``IOUtils.toByteArray``: its bytes, refused over `limit`.

        A negative size is no failure there, but what follows always is: the
        read takes the rest of the record, so the next read fails, and a string's
        conversion fails on the length.
        """
        if not 0 <= size <= limit:
            raise _NotRead(f"a length of {size}")
        return self.take(size)

    def ascii_z(self) -> str:
        """``readAsciiZ``: up to the first zero, within POI's 1,024 bytes; UTF-8, a malformed byte U+FFFD."""
        limit = max(self.at, min(self.at + _MAX_STRING, self._end))
        zero = self._data.find(0, self.at, limit)
        if zero < 0:
            raise _NotRead("a string with no terminator")
        out = self._data[self.at : zero]
        self.at = zero + 1
        # Java takes an encoded surrogate's three bytes as one malformed sequence, Python as three
        return _SURROGATE.sub("\ufffd".encode(), out).decode("utf-8", errors="replace")

    def ascii_len(self) -> str:
        """``readAsciiLen``: a length, then that many bytes, ISO-8859-1, the last (its terminator) dropped."""
        size = self.integer()
        return self.sized(size, _MAX_STRING)[: max(0, size - 1)].decode("latin-1")

    def utf16(self) -> None:
        """``readUtf16``, its text not kept: only that it can be read."""
        size = self.integer()
        self.sized(size * 2, _MAX_STRING)


def _control(byte: int) -> bool:
    """``Character.isISOControl`` of a signed byte: 0x80 and over are negative there, so never one."""
    return byte <= 0x1F or byte == 0x7F


def record_names(record: bytes) -> Names | None:
    """An ``Ole10Native`` record's label, file name and command, each None where empty; None for a record with none.

    Raises:
        Exception: where the record fails, which :func:`package_names` reads as none.
    """
    (total,) = struct.unpack_from("<i", record, 0)
    if total > _INT_MAX - 4:
        raise _NotRead("a total size past an int's end")  # POI's limit overflows, so that every read fails
    r = _Record(record, 4, min(total + 4, len(record)))
    if r.short() != 2:
        return None  # unparsed
    first = r.take(1)[0]
    r.at -= 1
    if _control(first):
        return None  # compact: an equation's record, say
    label = r.ascii_z()
    file_name = r.ascii_z()
    r.take(4)  # flags2, unknown1
    command = r.ascii_len()
    r.sized(r.integer(), _MAX_RECORD)  # the file
    mark = r.at
    if r.short() != 0:
        r.at = mark
        for _ in range(3):
            r.utf16()
    return label or None, file_name or None, command or None


def root_stream(data: bytes, name: str) -> bytes | None:
    """A stream of a compound file's root storage, its name matched without case; None for none.

    Raises:
        Exception: where `data` is no compound file this reader can read.
    """
    return _CompoundFile(data).stream(name)


def package_names(data: bytes) -> Names | None:
    """An OLE 1.0 package's label, file name and command; None where the bytes hold no ``Ole10Native`` stream.

    None too where the stream is no record, or fails anywhere (docx4j
    ``Embedded.ole10Names``). This never raises.
    """
    try:
        record = root_stream(data, OLE10_NATIVE)
        return None if record is None else record_names(record)
    except Exception:  # noqa: BLE001 - a file that gives no names gives none, whatever is wrong with it (CR-027 section 7)
        return None
