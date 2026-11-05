"""An OLE 1.0 package's bytes, for the text model's tests. CR-027 section 5.

Not a test. docx4j's ``EmbeddedObjectsTest.ole10`` builds its packages at test
time with POI (``Ole10Native.writeOut`` in a ``POIFSFileSystem``); here the
same bytes come from this writer, which needs no Java run. :func:`ole10_record`
is ``writeOut``'s parsed record, and :func:`compound_file` a version 3 or 4
compound file holding named streams in the root storage, those under 4,096
bytes in the mini stream, as POI writes them. The fixture ``probe41.bin``,
which POI wrote, is the reader's oracle beside it.
"""

from __future__ import annotations

import struct

__all__ = ["OLE_MARKER", "compound_file", "ole10", "ole10_record"]

#: The ``\x01Ole`` stream's content (POI ``Ole10Native.OLE_MARKER_BYTES``).
OLE_MARKER = bytes([1, 0, 0, 2]) + bytes(16)

_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_FAT_SECTOR = 0xFFFFFFFD
_DIFAT_SECTOR = 0xFFFFFFFC
_END = 0xFFFFFFFE
_FREE = 0xFFFFFFFF
_MINI_CUTOFF = 4096
_MINI_SECTOR = 64


def ole10_record(
    label: bytes | str = "notes.txt",
    file_name: bytes | str = "notes.txt",
    command: bytes | str = "notes.txt",
    content: bytes = b"NOTES_TEXT\r\n",
    *,
    flags1: int = 2,
    after: bytes = b"\0\0",
) -> bytes:
    """``Ole10Native.writeOut`` in its parsed form: `after` is what follows the data, by default no UTF-16 names."""

    def of(value: bytes | str) -> bytes:
        return value if isinstance(value, bytes) else value.encode()

    command_bytes = of(command) + b"\0"
    body = (
        struct.pack("<h", flags1)
        + of(label)
        + b"\0"
        + of(file_name)
        + b"\0"
        + struct.pack("<hh", 0, 3)
        + struct.pack("<i", len(command_bytes))
        + command_bytes
        + struct.pack("<i", len(content))
        + content
        + after
    )
    return struct.pack("<i", len(body)) + body


def _entry(name: str, kind: int, child: int, left: int, right: int, start: int, size: int) -> bytes:
    encoded = name.encode("utf-16-le")
    return (
        encoded.ljust(64, b"\0")
        + struct.pack("<HBB", len(encoded) + 2, kind, 1)
        + struct.pack("<III", left, right, child)
        + bytes(16 + 4 + 8 + 8)  # the class id, the state bits, the two times
        + struct.pack("<IQ", start, size)
    )


def compound_file(streams: dict[str, bytes], *, shift: int = 9, storages: tuple[str, ...] = ()) -> bytes:
    """A compound file: each stream in the root storage, a chain of right siblings; `storages` empty storages after them."""
    size = 1 << shift
    per_sector = size // 4
    mini = bytearray()
    mini_fat: list[int] = []
    big: list[bytes] = []  # what the FAT's chains hold, in sector order
    placed: dict[str, int] = {}
    for name, data in streams.items():
        if len(data) < _MINI_CUTOFF:
            count = -(-len(data) // _MINI_SECTOR)
            first = len(mini_fat)
            mini_fat += [first + i + 1 for i in range(count - 1)] + ([_END] if count else [])
            mini += data.ljust(count * _MINI_SECTOR, b"\0")
            placed[name] = first if count else _END
    sector = 0
    chains: list[int] = []  # the FAT

    def chain(data: bytes) -> int:
        nonlocal sector
        count = -(-len(data) // size)
        if not count:
            return _END
        first = sector
        chains.extend(first + i + 1 for i in range(count - 1))
        chains.append(_END)
        big.append(data.ljust(count * size, b"\0"))
        sector += count
        return first

    for name, data in streams.items():
        if name not in placed:
            placed[name] = chain(data)
    root_start = chain(bytes(mini))
    mini_fat_bytes = struct.pack(f"<{len(mini_fat)}I", *mini_fat).ljust(-(-len(mini_fat) * 4 // size) * size, b"\xff")
    first_mini_fat = chain(mini_fat_bytes)
    names = [*streams, *storages]
    entries = [_entry("Root Entry", 5, 1 if names else _FREE, _FREE, _FREE, root_start, len(mini))]
    for i, name in enumerate(names):
        following = i + 2 if i + 1 < len(names) else _FREE
        if name in streams:
            entries.append(_entry(name, 2, _FREE, _FREE, following, placed[name], len(streams[name])))
        else:
            entries.append(_entry(name, 1, _FREE, _FREE, following, _END, 0))
    directory = b"".join(entries).ljust(-(-len(entries) * 128 // size) * size, b"\0")
    first_directory = chain(directory)
    # the FAT's own sectors, and the DIFAT's beyond the header's 109, follow the data
    fat_count = difat_count = 0
    while True:
        total = sector + fat_count + difat_count
        needed = -(-total // per_sector)
        difat_needed = -(-max(0, needed - 109) // (per_sector - 1))
        if (needed, difat_needed) == (fat_count, difat_count):
            break
        fat_count, difat_count = needed, difat_needed
    fat_sectors = list(range(sector, sector + fat_count))
    difat_sectors = list(range(sector + fat_count, sector + fat_count + difat_count))
    fat = chains + [_FAT_SECTOR] * fat_count + [_DIFAT_SECTOR] * difat_count
    fat += [_FREE] * (fat_count * per_sector - len(fat))
    difat_blocks = []
    rest = fat_sectors[109:]
    for i in range(difat_count):
        taken, rest = rest[: per_sector - 1], rest[per_sector - 1 :]
        following = difat_sectors[i + 1] if i + 1 < difat_count else _END
        difat_blocks.append(struct.pack(f"<{per_sector}I", *taken, *[_FREE] * (per_sector - 1 - len(taken)), following))
    header = (
        _SIGNATURE
        + bytes(16)
        + struct.pack("<HHHHH", 0x3E, 3 if shift == 9 else 4, 0xFFFE, shift, 6)
        + bytes(6)
        + struct.pack("<II", 0 if shift == 9 else -(-len(directory) // size), fat_count)
        + struct.pack("<III", first_directory, 0, _MINI_CUTOFF)
        + struct.pack("<II", first_mini_fat, len(mini_fat_bytes) // size)
        + struct.pack("<II", difat_sectors[0] if difat_sectors else _END, difat_count)
        + struct.pack("<109I", *fat_sectors[:109], *[_FREE] * (109 - min(109, fat_count)))
    )
    return header.ljust(size, b"\0") + b"".join(big) + struct.pack(f"<{len(fat)}I", *fat) + b"".join(difat_blocks)


def ole10(label: str, file_name: str, command: str, content: str, *, shift: int = 9) -> bytes:
    """docx4j's ``ole10``: an OLE 1.0 package holding a text file, as Word embeds one."""
    record = ole10_record(label, file_name, command, content.encode())
    return compound_file({"\x01Ole": OLE_MARKER, "\x01Ole10Native": record}, shift=shift)
