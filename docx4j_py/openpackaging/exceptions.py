"""The engine's exceptions, as ``org.docx4j.openpackaging.exceptions``.

Four, and the hierarchy is docx4j's: everything the engine raises deliberately
is a :class:`Docx4JException`, so ``except Docx4JException`` catches the lot.
"""

from __future__ import annotations

__all__ = [
    "Docx4JException",
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


class InvalidOperationException(Docx4JException):
    """The operation is not allowed on this package. docx4j ``InvalidOperationException``."""


class PartUnrecognisedException(Docx4JException):
    """No part class is known for a content type. docx4j ``PartUnrecognisedException``."""
