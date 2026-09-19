"""The engine's exceptions, as ``org.docx4j.openpackaging.exceptions``.

Five, and the hierarchy is docx4j's: everything the engine raises deliberately
is a :class:`Docx4JException`, so ``except Docx4JException`` catches the lot.
"""

from __future__ import annotations

__all__ = [
    "CyclicStylesException",
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


class CyclicStylesException(Docx4JException):
    """A ``w:basedOn`` chain that comes back to a style it has already seen.

    docx4j ``CyclicStylesException``. Property resolution raises it only when
    :func:`docx4j_py.model.style_util.throw_on_cyclic_styles` has been turned on
    (docx4j's ``docx4j.openpackaging.exceptions.CyclicStylesException.throw``
    property, false by default); otherwise the walk stops where the cycle was
    found, one warning is logged, and resolution uses as much of the hierarchy
    as it has.
    """
