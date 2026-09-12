"""Where a package's bytes come from and go to. CR-002 section 4.

docx4j's ``io3.stores.PartStore`` separates the container from the package, and
that split is the point here because there are three containers from the start:
a zip, a directory and memory, with flat OPC to come in Phase C.

Names in a store are as the container stores them --- ``word/document.xml``,
``[Content_Types].xml``, ``_rels/.rels`` --- with **no leading slash**, which is
:attr:`PartName.store_name`. Lookup is case-insensitive, because
:class:`~docx4j_py.openpackaging.part_name.PartName` is and a store that was not
would put the two out of step.

| class | source | docx4j |
|---|---|---|
| :class:`ZipPartStore` | a path, bytes or a binary file object | ``ZipPartStore`` |
| :class:`DirectoryPartStore` | an unzipped directory | ``UnzippedPartStore`` |
| :class:`MemoryPartStore` | a ``dict`` | --- |
| :class:`ZipPartSink` | a path, a file object, or bytes | ``ZipPartStore``, the save half |
| :class:`DirectoryPartSink` | a directory | ``UnzippedPartStore`` |
| :class:`MemoryPartSink` | | --- |
"""

from __future__ import annotations

import os
import shutil
import zipfile
from collections.abc import Iterable, Iterator
from pathlib import Path
from typing import IO, Any, Protocol, runtime_checkable

from docx4j_py.openpackaging.exceptions import Docx4JException

__all__ = [
    "DirectoryPartSink",
    "DirectoryPartStore",
    "MemoryPartSink",
    "MemoryPartStore",
    "PartSink",
    "PartStore",
    "ZipPartSink",
    "ZipPartStore",
]


def _normalise(name: str) -> str:
    """A store key: no leading slash, backslashes folded to slashes."""
    name = name.replace("\\", "/")
    return name.removeprefix("/")


@runtime_checkable
class PartStore(Protocol):
    """Where a package's bytes come from. docx4j ``io3.stores.PartStore``, load half."""

    def part_names(self) -> Iterable[str]:
        """Every name the container holds, in container order, as stored."""
        ...

    def has(self, part_name: str) -> bool:
        """Whether the container holds a part of this name (case-insensitively)."""
        ...

    def load(self, part_name: str) -> bytes:
        """The bytes of a part. May inflate lazily on first access."""
        ...

    def size(self, part_name: str) -> int | None:
        """The uncompressed size if it is known without loading, else None."""
        ...


@runtime_checkable
class PartSink(Protocol):
    """Where a package's bytes go. docx4j ``PartStore``, the save half."""

    def put(
        self,
        part_name: str,
        data: bytes,
        *,
        content_type: str | None = None,
        compress: bool = True,
    ) -> None:
        """Write one part."""
        ...

    def finish(self) -> Any:
        """Complete the container and return whatever it produces."""
        ...


class _CaseInsensitiveIndex:
    """Lower-cased name -> the name as stored. The shared half of every store."""

    __slots__ = ("_index", "_order")

    def __init__(self, names: Iterable[str] = ()) -> None:
        self._index: dict[str, str] = {}
        self._order: list[str] = []
        for name in names:
            self.add(name)

    def add(self, name: str) -> str:
        stored = _normalise(name)
        key = stored.lower()
        if key not in self._index:
            self._order.append(stored)
        self._index[key] = stored
        return stored

    def resolve(self, name: str) -> str | None:
        """The stored spelling of a name, trying percent-decoding as docx4j does."""
        stored = _normalise(name)
        found = self._index.get(stored.lower())
        if found is not None:
            return found
        from urllib.parse import unquote

        decoded = unquote(stored)
        if decoded != stored:
            return self._index.get(decoded.lower())
        return None

    def remove(self, name: str) -> bool:
        stored = self.resolve(name)
        if stored is None:
            return False
        del self._index[stored.lower()]
        self._order.remove(stored)
        return True

    def names(self) -> list[str]:
        return list(self._order)

    def __contains__(self, name: str) -> bool:
        return self.resolve(name) is not None

    def __len__(self) -> int:
        return len(self._order)


# ---------------------------------------------------------------------------
# stores
# ---------------------------------------------------------------------------


class ZipPartStore:
    """A zip container, read lazily. docx4j ``ZipPartStore``.

    The central directory is read once at construction; an entry is inflated on
    the first :meth:`load` and is **not** cached, because the whole point of the
    lazy store is that a package which touches three parts does not pay to
    inflate thirty. The file stays open until :meth:`close`, which
    :meth:`~docx4j_py.openpackaging.packages.opc_package.OpcPackage.close` calls
    and which the package's context manager guarantees.

    ZIP64 needs no special handling: :mod:`zipfile` reads it.
    """

    __slots__ = ("_index", "_owns_file", "_path", "_source", "_zip")

    def __init__(self, source: str | os.PathLike[str] | bytes | IO[bytes]) -> None:
        """Open a zip from a path, a ``bytes`` object or a binary file object."""
        self._source = source
        self._path: Path | None = None
        self._owns_file = False
        if isinstance(source, (str, os.PathLike)):
            self._path = Path(source)
            self._owns_file = True
            handle: Any = str(self._path)
        elif isinstance(source, (bytes, bytearray, memoryview)):
            import io

            handle = io.BytesIO(bytes(source))
            self._owns_file = True
        else:
            handle = source
        try:
            self._zip = zipfile.ZipFile(handle)
        except zipfile.BadZipFile as exc:
            raise Docx4JException(f"Not a zip container: {exc}") from exc
        self._index = _CaseInsensitiveIndex(info.filename for info in self._zip.infolist())

    @property
    def path(self) -> Path | None:
        """The file this store was opened from, when it was opened from a path."""
        return self._path

    def part_names(self) -> list[str]:
        """Every entry, in central-directory order."""
        return self._index.names()

    def has(self, part_name: str) -> bool:
        """Whether the zip holds this entry."""
        return part_name in self._index

    def load(self, part_name: str) -> bytes:
        """Inflate one entry."""
        stored = self._index.resolve(part_name)
        if stored is None:
            raise Docx4JException(f"No part {part_name} in this zip")
        return self._zip.read(stored)

    def size(self, part_name: str) -> int | None:
        """The uncompressed size from the central directory."""
        stored = self._index.resolve(part_name)
        if stored is None:
            return None
        return self._zip.getinfo(stored).file_size

    def close(self) -> None:
        """Close the zip, and the file it owns."""
        self._zip.close()

    def reopen(self) -> None:
        """Reopen the same path, after saving over it (decided question 4)."""
        if self._path is None:
            raise Docx4JException("This store was not opened from a path")
        self._zip.close()
        self._zip = zipfile.ZipFile(str(self._path))
        self._index = _CaseInsensitiveIndex(info.filename for info in self._zip.infolist())

    def __repr__(self) -> str:
        """``ZipPartStore(18 parts)``."""
        return f"ZipPartStore({len(self._index)} parts)"


class DirectoryPartStore:
    """An unzipped package. docx4j ``UnzippedPartStore``; for diffing and tests."""

    __slots__ = ("_index", "_root")

    def __init__(self, root: str | os.PathLike[str]) -> None:
        """Index every file under `root`, recursively."""
        self._root = Path(root)
        if not self._root.is_dir():
            raise Docx4JException(f"{self._root} is not a directory")
        names = []
        for path in sorted(self._root.rglob("*")):
            if path.is_file():
                names.append(path.relative_to(self._root).as_posix())
        self._index = _CaseInsensitiveIndex(names)

    @property
    def root(self) -> Path:
        """The directory."""
        return self._root

    def part_names(self) -> list[str]:
        """Every file under the directory, in sorted order."""
        return self._index.names()

    def has(self, part_name: str) -> bool:
        """Whether the directory holds this file."""
        return part_name in self._index

    def load(self, part_name: str) -> bytes:
        """Read one file."""
        stored = self._index.resolve(part_name)
        if stored is None:
            raise Docx4JException(f"No part {part_name} in {self._root}")
        return (self._root / stored).read_bytes()

    def size(self, part_name: str) -> int | None:
        """The file's size."""
        stored = self._index.resolve(part_name)
        if stored is None:
            return None
        return (self._root / stored).stat().st_size

    def close(self) -> None:
        """Nothing to close; here so every store answers the same calls."""

    def __repr__(self) -> str:
        """``DirectoryPartStore('/tmp/x', 18 parts)``."""
        return f"DirectoryPartStore({str(self._root)!r}, {len(self._index)} parts)"


class MemoryPartStore:
    """A package in a ``dict``: new packages, tests, and a sink's result."""

    __slots__ = ("_index", "_parts", "content_types")

    def __init__(self, parts: dict[str, bytes] | None = None) -> None:
        """Build a store, optionally filled from a mapping of name to bytes."""
        self._parts: dict[str, bytes] = {}
        self._index = _CaseInsensitiveIndex()
        #: Content types as they were put, for a store filled through a sink.
        self.content_types: dict[str, str] = {}
        for name, data in (parts or {}).items():
            self.put(name, data)

    def part_names(self) -> list[str]:
        """Every name, in insertion order."""
        return self._index.names()

    def has(self, part_name: str) -> bool:
        """Whether the store holds this part."""
        return part_name in self._index

    def load(self, part_name: str) -> bytes:
        """The bytes of a part."""
        stored = self._index.resolve(part_name)
        if stored is None:
            raise Docx4JException(f"No part {part_name} in this store")
        return self._parts[stored]

    def size(self, part_name: str) -> int | None:
        """The length of a part."""
        stored = self._index.resolve(part_name)
        return None if stored is None else len(self._parts[stored])

    def put(
        self,
        part_name: str,
        data: bytes,
        *,
        content_type: str | None = None,
        compress: bool = True,  # noqa: ARG002 - memory does not compress
    ) -> None:
        """Store a part; the store is its own sink."""
        stored = self._index.add(part_name)
        self._parts[stored] = bytes(data)
        if content_type is not None:
            self.content_types[stored] = content_type

    def delete(self, part_name: str) -> bool:
        """Remove a part; True when there was one."""
        stored = self._index.resolve(part_name)
        if stored is None:
            return False
        self._index.remove(stored)
        del self._parts[stored]
        return True

    def finish(self) -> MemoryPartStore:
        """Itself; the store is its own sink."""
        return self

    def close(self) -> None:
        """Nothing to close."""

    def __repr__(self) -> str:
        """``MemoryPartStore(18 parts)``."""
        return f"MemoryPartStore({len(self._index)} parts)"


# ---------------------------------------------------------------------------
# sinks
# ---------------------------------------------------------------------------


class ZipPartSink:
    """Write a zip. docx4j ``ZipPartStore``, the save half.

    ``[Content_Types].xml`` goes in first because the caller writes it first
    (docx4j ``Save.save``), which is the order Word writes and the order some
    readers assume. Already-compressed media is **stored** rather than deflated,
    as ``ZipPartStore.shouldCompress`` decides; everything else is deflated.
    """

    __slots__ = ("_names", "_own_stream", "_target", "_zip")

    def __init__(
        self,
        target: str | os.PathLike[str] | IO[bytes] | None = None,
        *,
        compresslevel: int | None = None,
    ) -> None:
        """Write to a path, to a binary file object, or (with None) to bytes."""
        self._own_stream = None
        self._names: set[str] = set()
        if target is None:
            import io

            self._own_stream = io.BytesIO()
            self._target: Any = self._own_stream
        elif isinstance(target, (str, os.PathLike)):
            self._target = str(target)
        else:
            self._target = target
        kwargs: dict[str, Any] = {"allowZip64": True}
        if compresslevel is not None:
            kwargs["compresslevel"] = compresslevel
        self._zip = zipfile.ZipFile(self._target, "w", zipfile.ZIP_DEFLATED, **kwargs)

    def put(
        self,
        part_name: str,
        data: bytes,
        *,
        content_type: str | None = None,  # noqa: ARG002 - a zip carries it in [Content_Types].xml
        compress: bool = True,
    ) -> None:
        """Write one entry; `compress` False stores it."""
        name = _normalise(part_name)
        if name in self._names:
            return
        self._names.add(name)
        method = zipfile.ZIP_DEFLATED if compress else zipfile.ZIP_STORED
        info = zipfile.ZipInfo(name)
        info.compress_type = method
        info.external_attr = 0o600 << 16
        self._zip.writestr(info, data)

    def finish(self) -> bytes | None:
        """Close the zip; the bytes when the sink was built with no target."""
        self._zip.close()
        if self._own_stream is not None:
            return self._own_stream.getvalue()
        return None

    def __repr__(self) -> str:
        """``ZipPartSink(18 entries)``."""
        return f"ZipPartSink({len(self._names)} entries)"


class DirectoryPartSink:
    """Write an unzipped package, for diffing saved output."""

    __slots__ = ("_names", "_root")

    def __init__(self, root: str | os.PathLike[str], *, clean: bool = False) -> None:
        """Write under `root`, creating it; `clean` empties it first."""
        self._root = Path(root)
        if clean and self._root.exists():
            shutil.rmtree(self._root)
        self._root.mkdir(parents=True, exist_ok=True)
        self._names: list[str] = []

    def put(
        self,
        part_name: str,
        data: bytes,
        *,
        content_type: str | None = None,  # noqa: ARG002
        compress: bool = True,  # noqa: ARG002 - a directory does not compress
    ) -> None:
        """Write one file, creating its directories."""
        name = _normalise(part_name)
        path = self._root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        self._names.append(name)

    def finish(self) -> Path:
        """The directory."""
        return self._root

    def __repr__(self) -> str:
        """``DirectoryPartSink('/tmp/x', 18 files)``."""
        return f"DirectoryPartSink({str(self._root)!r}, {len(self._names)} files)"


class MemoryPartSink:
    """Collect into a fresh :class:`MemoryPartStore`."""

    __slots__ = ("store",)

    def __init__(self) -> None:
        """Build the sink and the store behind it."""
        self.store = MemoryPartStore()

    def put(
        self,
        part_name: str,
        data: bytes,
        *,
        content_type: str | None = None,
        compress: bool = True,
    ) -> None:
        """Store one part."""
        self.store.put(part_name, data, content_type=content_type, compress=compress)

    def finish(self) -> MemoryPartStore:
        """The store."""
        return self.store

    def __repr__(self) -> str:
        """``MemoryPartSink(18 parts)``."""
        return f"MemoryPartSink({len(self.store.part_names())} parts)"


def store_for(source: Any) -> PartStore:
    """The right store for a path, a directory, bytes, a file object or a store.

    What :meth:`OpcPackage.load` calls to make sense of its argument.
    """
    if isinstance(source, (ZipPartStore, DirectoryPartStore, MemoryPartStore)):
        return source
    if isinstance(source, (str, os.PathLike)):
        path = Path(source)
        if path.is_dir():
            return DirectoryPartStore(path)
        return ZipPartStore(path)
    if isinstance(source, (bytes, bytearray, memoryview)):
        return ZipPartStore(bytes(source))
    if hasattr(source, "read"):
        return ZipPartStore(source)
    if isinstance(source, PartStore):
        return source
    raise Docx4JException(f"Cannot load a package from {type(source).__name__}")


def iter_store(store: PartStore) -> Iterator[tuple[str, bytes]]:
    """Every part of a store as ``(name, bytes)``; for tests and diffs."""
    for name in store.part_names():
        yield name, store.load(name)
