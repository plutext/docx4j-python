"""What a field evaluator may ask for. CR-025 Phase B, D5 and D10.

docx4j's ``org.docx4j.model.fields.FieldContext`` at ``c3e9920e5`` (its CR-067
D10, D12): the package; the zone a date field's instant is shown in (this
context's, else the package's ``date_time_zone``, else the process's); a
clock (``DATE``, ``TIME``, ``PRINTDATE``); the language a result is formatted
in (None, as docx4j's inspector passes none); the user (``USERNAME``,
``USERINITIALS``, ``USERADDRESS``), absent by default; the section the field
is in (``SECTION``); and the state one pass over a document accumulates
(``SEQ`` counters, the text of the bookmarks seen so far, the note numbers).
What is not supplied is absent, and a field that needs it keeps its cached
result.

One context serves one pass over one document. Not docx4j's: the file name and
the external resolver (no evaluator here reads them), the part, and
``section_of`` (the consumer here, the text model, counts its own sections).
:attr:`FieldContext.computes` is this port's: the names the consumer
evaluates, which ``INFO`` asks before it evaluates the field it names (CR-025
D12).
"""

from __future__ import annotations

import dataclasses
import datetime
from collections.abc import Callable, Collection, Mapping
from typing import Any

__all__ = ["FieldContext", "FieldUser"]


@dataclasses.dataclass(frozen=True, slots=True)
class FieldUser:
    """The user ``USERNAME``, ``USERINITIALS`` and ``USERADDRESS`` show. docx4j ``FieldContext.user(name, initials, address)``.

    Word shows its own user (File > Options > General); a document has none,
    so the caller supplies one, or the three fields keep their cached result.
    Initials or an address left None show empty, as Word shows a blank one
    (docx4j CR-067 section 11.3, case 22).
    """

    name: str
    initials: str | None = None
    address: str | None = None

    @classmethod
    def of(cls, value: Any) -> FieldUser | None:
        """A user from what a caller gives: None, a :class:`FieldUser`, a name, or a mapping of the three.

        Raises:
            TypeError: anything else, or a mapping with another key or no name.
        """
        if value is None or isinstance(value, FieldUser):
            return value
        if isinstance(value, str):
            return cls(value)
        if isinstance(value, Mapping):
            extra = set(value) - {"name", "initials", "address"}
            if extra or not isinstance(value.get("name"), str):
                raise TypeError(f"a field user is a mapping of name, initials and address, not {dict(value)!r}")
            return cls(value["name"], value.get("initials"), value.get("address"))
        raise TypeError(f"a field user is a FieldUser, a name or a mapping, not {type(value).__name__}")


def _now() -> datetime.datetime:
    return datetime.datetime.now(datetime.UTC)


class FieldContext:
    """One pass's context: the package, the zone, the clock, the user and what the pass has counted."""

    __slots__ = (
        "_note_numbers",
        "_properties",
        "_variables",
        "_zone",
        "bookmarks",
        "clock",
        "computes",
        "document",
        "language",
        "package",
        "section",
        "sequences",
        "user",
    )

    def __init__(
        self,
        package: Any = None,
        *,
        zone: Any = None,
        clock: Callable[[], datetime.datetime] | None = None,
        language: str | None = None,
        user: Any = None,
        section: int | None = None,
        document: Any = None,
        computes: Collection[str] | None = None,
    ) -> None:
        """A context for `package`, with the defaults; each keyword is the attribute of its name.

        Raises:
            ValueError: a `zone` name the platform does not know.
            TypeError: a `zone` or a `user` of another type.
        """
        from docx4j_py.model.dates import zone_of

        #: The package the fields are in.
        self.package = package
        self._zone = zone_of(zone)
        #: What gives "now" (``DATE``, ``TIME``, ``PRINTDATE``): an aware datetime; now by default.
        self.clock: Callable[[], datetime.datetime] = clock or _now
        #: The language a result is formatted in, or None: US English, as docx4j's inspector passes none.
        self.language = language.strip() if isinstance(language, str) and language.strip() else None
        #: The user the user fields show, or None (they keep their cached result).
        self.user: FieldUser | None = FieldUser.of(user)
        #: The section (from 1) of the main story the field is in, or None (``SECTION`` keeps its cached result).
        self.section = section
        #: The ``w:document`` the note numbers and bookmarks are read from, else the package's main part's.
        self.document = document
        #: The names the consumer evaluates, upper case, or None for every name registered.
        self.computes: frozenset[str] | None = None if computes is None else frozenset(n.upper() for n in computes)
        #: ``SEQ`` counters by identifier (upper case), this pass's.
        self.sequences: dict[str, int] = {}
        #: The text of the bookmarks this pass has seen, by name in lower case; a ``REF`` reads it.
        self.bookmarks: dict[str, str] = {}
        self._note_numbers: dict[str, str] | None = None
        self._properties: Any = None
        self._variables: Any = None

    @property
    def zone(self) -> datetime.tzinfo:
        """The zone a date field's instant is shown in: this context's, else the package's, else the process's.

        docx4j ``FieldContext.getZone()`` (D12). The process's zone is
        :data:`docx4j_py.model.dates.PROCESS_ZONE`.
        """
        from docx4j_py.model.dates import effective_zone

        return self._zone if self._zone is not None else effective_zone(self.package)

    @zone.setter
    def zone(self, value: Any) -> None:
        """The per-call override (docx4j ``FieldContext.zone(ZoneId)``); None returns to the package's."""
        from docx4j_py.model.dates import zone_of

        self._zone = zone_of(value)

    def now(self) -> datetime.datetime:
        """The moment ``DATE``, ``TIME`` and ``PRINTDATE`` show: the clock's, aware (a naive one is UTC)."""
        moment = self.clock()
        return moment if moment.tzinfo is not None else moment.replace(tzinfo=datetime.UTC)

    def evaluates(self, name: str | None) -> bool:
        """Whether the consumer evaluates a field of this name (:attr:`computes`; every name where it is None)."""
        return name is not None and (self.computes is None or name.upper() in self.computes)

    def _document(self) -> Any:
        if self.document is not None:
            return self.document
        main = getattr(self.package, "main_document_part", None) if self.package is not None else None
        if main is None:
            return None
        from docx4j_py.model.fields.resolvers import _contents

        return _contents(main)

    def note_numbers(self) -> dict[str, str]:
        """The note number each bookmark holds a reference to, by name in lower case (``NOTEREF``), read once."""
        if self._note_numbers is None:
            from docx4j_py.model.fields.sources import note_numbers

            document = self._document()
            self._note_numbers = note_numbers(document) if document is not None else {}
        return self._note_numbers

    def read_bookmarks(self) -> FieldContext:
        """Fill :attr:`bookmarks` from the main document, for a caller with no reading of its own; returns self."""
        if not self.bookmarks:
            from docx4j_py.model.fields.sources import bookmark_texts

            document = self._document()
            if document is not None:
                self.bookmarks.update(bookmark_texts(document))
        return self

    def properties(self) -> Any:
        """The document's properties (core, application, custom), read once."""
        if self._properties is None:
            from docx4j_py.model.fields.resolvers import DocPropertyResolver

            self._properties = DocPropertyResolver(self.package)
        return self._properties

    def variables(self) -> Any:
        """The document's variables, read once."""
        if self._variables is None:
            from docx4j_py.model.fields.resolvers import DocVariableResolver

            self._variables = DocVariableResolver(self.package)
        return self._variables
