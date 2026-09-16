"""``DocumentSession``: many documents open across many tool calls.

CR-003 section 3.4 and decided question 5. CR-002 accepted a one-second import
on the condition that a process serves many documents; an MCP server therefore
holds packages open across tool calls, so that "open, edit, edit, save" is four
cheap calls and not four loads. This is in the library rather than in the server
because the thread rule and the idle-close logic are the library's knowledge and
every server would otherwise write them::

    session = DocumentSession(idle_timeout=900, max_open=32)

    handle = session.open("report.docx")          # a short opaque string
    with session.use(handle) as pkg:              # the handle's lock is held
        pkg.body.paragraph_at(contains="Chapter 1").insert_paragraph("New", location="After")
        report = pkg.last_change
    session.save(handle, "edited.docx")           # or save(handle) for the bytes,
    session.close(handle)                         # or overwrite=True for the source

**One lock per handle.** Two tool calls that arrive at once for the same
document serialise; two for different documents do not. :meth:`use` is the way
to hold it for a whole edit; :meth:`get`, :meth:`save` and :meth:`close` take it
for themselves.

**One ``ParserConfig`` per thread**, as CR-001 section 14.7 requires: nothing
here has to arrange it, because
:class:`~docx4j_py.openpackaging.parts.xml_part.XmlPart` already takes a fresh
parser --- and therefore a fresh config, and therefore its own skipped-content
report --- for **every** unmarshal
(:func:`docx4j_py.runtime.parser` returns a new one per call). The one process-wide
thing, the ``XmlContext``, is safe to share and is what makes
:func:`docx4j_py.wml.warm_up` worth anything across workers. A test asserts
both halves.

**No background thread.** :attr:`idle_timeout` is enforced by :meth:`sweep`,
which the server calls --- between requests, on a timer of its own, or never.
A library that starts threads in a server's process is a library that is hard
to embed.
"""

from __future__ import annotations

import os
import secrets
import threading
import time
from collections.abc import Iterator
from contextlib import contextmanager
from typing import IO, Any, Self

from docx4j_py.model.content.errors import ContentError

__all__ = ["DocumentSession", "OpenDocument"]

#: How long a handle is, in bytes of entropy; ``token_urlsafe`` makes it 8 characters.
_HANDLE_BYTES = 6

#: The default idle timeout, in seconds. Fifteen minutes, which is long enough
#: for a conversation and short enough that a forgotten document is reclaimed.
DEFAULT_IDLE_TIMEOUT = 900.0

#: How many documents one session holds before :meth:`DocumentSession.open` refuses.
DEFAULT_MAX_OPEN = 32


class OpenDocument:
    """One open document: the package, its lock, and when it was last used."""

    __slots__ = ("handle", "last_used", "lock", "opened_at", "package", "source")

    def __init__(self, handle: str, package: Any, source: Any) -> None:
        """Hold an open package under a handle."""
        #: The handle it is reached by.
        self.handle = handle
        #: The :class:`~docx4j_py.openpackaging.packages.opc_package.OpcPackage`.
        self.package = package
        #: What it was opened from: a path, or None for bytes and file objects.
        self.source = source
        #: The lock that serialises tool calls on this document.
        self.lock = threading.RLock()
        #: When it was opened (``time.monotonic``).
        self.opened_at = time.monotonic()
        #: When it was last used, which is what :meth:`DocumentSession.sweep` reads.
        self.last_used = self.opened_at

    @property
    def idle_for(self) -> float:
        """Seconds since this document was last used."""
        return time.monotonic() - self.last_used

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: what a ``list_open`` tool returns."""
        return {
            "handle": self.handle,
            "source": str(self.source) if self.source is not None else None,
            "idle_for": round(self.idle_for, 3),
            "changes": len(getattr(self.package, "changes", ()) or ()),
        }

    def __repr__(self) -> str:
        """``<OpenDocument 3x_Kd9Qa report.docx, idle 4.2s>``."""
        return f"<OpenDocument {self.handle} {self.source}, idle {self.idle_for:.1f}s>"


class DocumentSession:
    """A registry of open packages, keyed by a short opaque handle."""

    __slots__ = ("_documents", "_lock", "idle_timeout", "max_open")

    def __init__(
        self,
        *,
        idle_timeout: float = DEFAULT_IDLE_TIMEOUT,
        max_open: int = DEFAULT_MAX_OPEN,
    ) -> None:
        """Build an empty session.

        Args:
            idle_timeout: seconds of inactivity after which :meth:`sweep`
                closes a document. ``0`` or less turns it off.
            max_open: how many documents may be open at once; :meth:`open`
                sweeps first and then refuses.
        """
        #: Seconds of inactivity :meth:`sweep` closes a document after.
        self.idle_timeout = idle_timeout
        #: How many documents may be open at once.
        self.max_open = max_open
        self._documents: dict[str, OpenDocument] = {}
        self._lock = threading.RLock()

    # -- opening and closing ----------------------------------------------

    def open(
        self,
        source: str | os.PathLike[str] | bytes | IO[bytes],
        *,
        options: Any = None,
    ) -> str:
        """Load a document and return the handle it is reached by.

        Args:
            source: anything :func:`docx4j_py.load` takes --- a path, a
                directory, bytes, a file object or a ``PartStore``.
            options: a :class:`~docx4j_py.openpackaging.load.LoadOptions`.

        Returns:
            A short opaque string. Opaque on purpose: it is not a path, so it
            carries nothing a tool result should not carry, and a server that
            returns it is not returning the caller's file system.

        Raises:
            ContentError: the session is full; the message says to close one.
        """
        from docx4j_py.openpackaging.api import load

        with self._lock:
            if len(self._documents) >= self.max_open:
                self.sweep()
            if len(self._documents) >= self.max_open:
                raise ContentError(
                    f"this session already holds {len(self._documents)} documents, "
                    f"which is max_open",
                    code="session.full",
                    hint="close(handle) one you have finished with, or raise max_open",
                )
        package = load(source, options=options)
        keep = source if isinstance(source, (str, os.PathLike)) else None
        with self._lock:
            handle = self._new_handle()
            self._documents[handle] = OpenDocument(handle, package, keep)
            return handle

    def add(self, package: Any, *, source: Any = None) -> str:
        """Put an already-open package in the session. Extension.

        For ``create_package()``, and for a server that loaded a document some
        other way. `source` is what :meth:`save` writes to when it is given no
        target.
        """
        with self._lock:
            handle = self._new_handle()
            self._documents[handle] = OpenDocument(handle, package, source)
            return handle

    def close(self, handle: str) -> None:
        """Close a document and forget its handle. Idempotent."""
        with self._lock:
            document = self._documents.pop(handle, None)
        if document is None:
            return
        with document.lock:
            close = getattr(document.package, "close", None)
            if close is not None:
                close()

    def close_all(self) -> None:
        """Close every document. What ``__exit__`` does."""
        for handle in self.handles():
            self.close(handle)

    # -- using -------------------------------------------------------------

    def get(self, handle: str) -> Any:
        """The package behind a handle, and mark it used.

        The handle's lock is **not** held afterwards: use :meth:`use` for an
        edit that has to be atomic against another tool call.

        Raises:
            ContentError: no such handle, or it has been closed or swept.
        """
        document = self._document(handle)
        document.last_used = time.monotonic()
        return document.package

    @contextmanager
    def use(self, handle: str) -> Iterator[Any]:
        """``with session.use(handle) as pkg:`` --- the package, lock held.

        Two tool calls on one document serialise here; two on different
        documents do not wait for each other.
        """
        document = self._document(handle)
        with document.lock:
            document.last_used = time.monotonic()
            try:
                yield document.package
            finally:
                document.last_used = time.monotonic()

    def lock_for(self, handle: str) -> threading.RLock:
        """The handle's lock, for a caller who wants to hold it itself."""
        return self._document(handle).lock

    def save(
        self,
        handle: str,
        target: str | os.PathLike[str] | IO[bytes] | None = None,
        *,
        overwrite: bool = False,
    ) -> bytes | None:
        """Save a document, holding its lock.

        Args:
            target: where to write: a path, a file object, or None for the
                bytes. Writing over the file the document was opened from is
                ``target=None, overwrite=True``, which is docx4j-mcp's
                ``overwrite: true``: a server should not silently rewrite a
                user's file because a tool call left the target out.
            overwrite: write over what the document was opened from.

        Returns:
            The bytes when nothing was written to a file, else None.

        Raises:
            ContentError: `overwrite` was asked for and the document was not
                opened from a path.
        """
        document = self._document(handle)
        if overwrite and document.source is None:
            raise ContentError(
                "this document was not opened from a path, so there is nothing to overwrite",
                code="session.no_source",
                hint="pass a target path, or save(handle) for the bytes",
            )
        where = document.source if overwrite else target
        with document.lock:
            document.last_used = time.monotonic()
            return document.package.save(where)

    # -- the registry ------------------------------------------------------

    def handles(self) -> list[str]:
        """Every open handle, in the order the documents were opened."""
        with self._lock:
            return list(self._documents)

    def documents(self) -> list[OpenDocument]:
        """Every open document, for a ``list_open`` tool."""
        with self._lock:
            return list(self._documents.values())

    def sweep(self) -> list[str]:
        """Close everything idle longer than :attr:`idle_timeout`.

        There is no background thread: the server calls this, between requests
        or on a timer of its own. A document whose lock is held by another
        thread is left alone and swept next time.

        Returns:
            The handles closed.
        """
        if self.idle_timeout is None or self.idle_timeout <= 0:
            return []
        with self._lock:
            stale = [
                document
                for document in self._documents.values()
                if document.idle_for > self.idle_timeout
            ]
        closed: list[str] = []
        for document in stale:
            if not document.lock.acquire(blocking=False):
                continue
            try:
                if document.idle_for <= self.idle_timeout:
                    continue
                with self._lock:
                    self._documents.pop(document.handle, None)
                close = getattr(document.package, "close", None)
                if close is not None:
                    close()
                closed.append(document.handle)
            finally:
                document.lock.release()
        return closed

    # -- Python's shapes ---------------------------------------------------

    def __len__(self) -> int:
        """How many documents are open."""
        with self._lock:
            return len(self._documents)

    def __contains__(self, handle: object) -> bool:
        """Whether a handle is open."""
        with self._lock:
            return handle in self._documents

    def __enter__(self) -> Self:
        """The session."""
        return self

    def __exit__(self, *exc_info: object) -> None:
        """Close every document still open."""
        self.close_all()

    def __repr__(self) -> str:
        """``<DocumentSession 3 open, max 32, idle timeout 900s>``."""
        return (
            f"<DocumentSession {len(self)} open, max {self.max_open}, "
            f"idle timeout {self.idle_timeout}s>"
        )

    # -- internals ---------------------------------------------------------

    def _new_handle(self) -> str:
        while True:
            handle = secrets.token_urlsafe(_HANDLE_BYTES)
            if handle not in self._documents:
                return handle

    def _document(self, handle: str) -> OpenDocument:
        with self._lock:
            document = self._documents.get(handle)
        if document is None:
            raise ContentError(
                f"no document open under the handle {handle!r}",
                code="session.unknown_handle",
                hint="call open() for a handle, or handles() for the ones still open; "
                "an idle document may have been swept",
            )
        return document
