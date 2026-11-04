"""The engine's exceptions, as ``org.docx4j.openpackaging.exceptions``.

Six, and the hierarchy is docx4j's: everything the engine raises deliberately
is a :class:`Docx4JException`, so ``except Docx4JException`` catches the lot.
:class:`DuplicateEntryException` is docx4j-core-ts's (its CR-024 D4), which
docx4j does not have yet.
"""

from __future__ import annotations

from collections.abc import Mapping
from typing import ClassVar

__all__ = [
    "CyclicStylesException",
    "Docx4JException",
    "DuplicateEntryException",
    "InvalidFormatException",
    "InvalidOperationException",
    "PartUnrecognisedException",
]


class Docx4JException(Exception):
    """Anything the engine refuses to do. docx4j ``Docx4JException``."""


class InvalidFormatException(Docx4JException):
    """The package, a part name or a relationship is not well formed.

    docx4j ``InvalidFormatException``. The OPC conformance clauses ([M1.1] and
    the rest) are quoted in the message, as docx4j quotes them.
    """


class DuplicateEntryException(InvalidFormatException):
    """A package whose container holds two entries of one name, refused (CR-032 D4, WD REQ-206).

    docx4j-core-ts ``DuplicateEntryException`` (its CR-024 D4). Raised by the
    load before it reads anything, under ``LoadOptions(duplicate_entries=
    "refuse")``, the default. Two names are one where the store's lookup
    would confuse them, as OPC part names compare: without case.

    Attributes:
        code: ``package.duplicate_entry``, the WD's code.
        entries: each duplicated name, as its first entry spells it
            (``word/document.xml``, ``[Content_Types].xml``), to its number of
            entries, in the container's order; so that a host can act on it
            without reading the message.
    """

    code: ClassVar[str] = "package.duplicate_entry"

    def __init__(self, entries: Mapping[str, int]) -> None:
        """The refusal, naming every duplicated name in one message."""
        self.entries: dict[str, int] = dict(entries)
        super().__init__(_duplicate_message(self.entries))


def _duplicate_message(entries: Mapping[str, int]) -> str:
    """``Duplicate entry word/document.xml in this package: 2 entries (REQ-206). LoadOptions(...) reads the first``."""
    option = 'LoadOptions(duplicate_entries="first") reads the first'
    if len(entries) == 1:
        ((name, count),) = entries.items()
        return f"Duplicate entry {name} in this package: {count} entries (REQ-206). {option}"
    listed = "; ".join(f"{name}, {count} entries" for name, count in entries.items())
    return f"Duplicate entries in this package: {listed} (REQ-206). {option} of each"


class InvalidOperationException(Docx4JException):
    """The operation is not allowed on this package. docx4j ``InvalidOperationException``."""


class PartUnrecognisedException(Docx4JException):
    """No part class is known for a content type. docx4j ``PartUnrecognisedException``."""


class CyclicStylesException(Docx4JException):
    """A ``w:basedOn`` chain that comes back to a style it has already seen.

    docx4j ``CyclicStylesException``. Property resolution raises it only when
    :func:`docx4j_py.model.style_util.throw_on_cyclic_styles` has been turned on
    (docx4j's ``docx4j.openpackaging.exceptions.CyclicStylesException.throw``
    property, false by default); otherwise the walk stops where the cycle was
    found, one warning is logged, and resolution uses as much of the hierarchy
    as it has.
    """
