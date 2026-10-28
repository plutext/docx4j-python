"""What the tier-1 evaluators read from the document itself. CR-025 Phase B.

docx4j's ``org.docx4j.model.fields.FieldSources`` at ``c3e9920e5`` (its CR-067
phase 1): the number of the note a bookmark holds a reference to
(``NOTEREF``), a bookmark's text (``REF``, for a caller with no reading of its
own), and a core-property date parsed (``CREATEDATE``, ``SAVEDATE``; a value
with no offset UTC since docx4j's D13, CR-025 Phase B2). Read
from the main document's tree; the text model's walk has its own reading of
the bookmarks and passes the tree it holds, so nothing here unmarshals a part.

The walk goes where docx4j's goes: the body's content lists in document order,
into paragraphs, runs, tables, hyperlinks, simple fields, content controls,
custom XML, smart tags and insertions; not into a deletion, a move, a drawing,
a picture, an object, an equation, ruby or ``mc:AlternateContent``, which
docx4j's walk does not enter either.
"""

from __future__ import annotations

import datetime
import re
from typing import Any

__all__ = ["bookmark_texts", "note_numbers", "parse_w3cdtf"]

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: What docx4j's walk does not enter (``FieldSources.walk``: no content accessor).
_NOT_ENTERED = frozenset(
    {
        _W + "del",
        _W + "moveFrom",
        _W + "moveTo",
        _W + "drawing",
        _W + "pict",
        _W + "object",
        _W + "ruby",
        "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath",
        "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara",
        "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent",
    }
)


class _Walk:
    """docx4j's ``FieldSources.Visitor`` over a ``w:document``'s body: bookmarks, text and note references."""

    def __init__(self) -> None:
        self.open: dict[Any, str] = {}

    def run(self, document: Any) -> None:
        from docx4j_py.traversal import walk

        body = getattr(document, "body", None)
        if body is None:
            return
        walk(body, self._visit, mce="none")

    def _visit(self, node: Any, parent: Any, name: str | None) -> bool | None:
        if name in _NOT_ENTERED:
            return False
        if name == _W + "bookmarkStart":
            bookmark = getattr(node, "name", None)
            if bookmark is not None:
                self.open[_key(getattr(node, "id", None))] = str(bookmark).lower()
                self.start(str(bookmark).lower())
            return False
        if name == _W + "bookmarkEnd":
            self.open.pop(_key(getattr(node, "id", None)), None)
            return False
        if name in (_W + "footnoteReference", _W + "endnoteReference"):
            self.note(name.rsplit("}", 1)[1], node)
            return False
        if name == _W + "t":
            self.text(getattr(node, "value", None) or "")
            return False
        return None

    def start(self, name: str) -> None:
        """A bookmark opens."""

    def text(self, value: str) -> None:
        """Text, in a ``w:t``."""

    def note(self, kind: str, ref: Any) -> None:
        """A footnote or endnote reference."""


def _key(value: Any) -> Any:
    from docx4j_py.model.values import integer

    found = integer(value)
    return found if found is not None else value


def note_numbers(document: Any) -> dict[str, str]:
    """The number of the footnote or endnote whose reference each bookmark holds, by bookmark name in lower case.

    docx4j ``FieldSources.noteNumbers``: references counted in document order
    per kind, from 1 (Word's default numbering); a reference with a custom
    mark (``w:customMarkFollows``) is not counted; one with an id of 0 or less
    (a separator's) is skipped; a bookmark holding several takes its first.
    """
    from docx4j_py.model.values import integer, on_off

    out: dict[str, str] = {}
    counts = {"footnoteReference": 0, "endnoteReference": 0}

    class Notes(_Walk):
        def note(self, kind: str, ref: Any) -> None:
            identifier = integer(getattr(ref, "id", None))
            if identifier is not None and identifier <= 0:
                return
            if not on_off(getattr(ref, "custom_mark_follows", None), default=False):
                counts[kind] += 1
            number = str(counts[kind])
            for name in self.open.values():
                out.setdefault(name, number)

    Notes().run(document)
    return out


def bookmark_texts(document: Any) -> dict[str, str]:
    """Each bookmark's text, its ``w:t`` in document order, by name in lower case. docx4j ``FieldSources.bookmarkTexts``."""
    texts: dict[str, list[str]] = {}

    class Texts(_Walk):
        def start(self, name: str) -> None:
            texts.setdefault(name, [])

        def text(self, value: str) -> None:
            for name in self.open.values():
                texts[name].append(value)

    Texts().run(document)
    return {name: "".join(parts) for name, parts in texts.items()}


_W3CDTF = re.compile(
    r"(\d{4})(?:-(\d{2})(?:-(\d{2})(?:[Tt](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?)?)?([Zz]|[+-]\d{2}:\d{2})?"
)


def parse_w3cdtf(value: str | None) -> datetime.datetime | None:
    """A core property's W3CDTF date as an instant, aware and in UTC; None for anything else.

    docx4j ``FieldSources.parseW3CDTF`` since its CR-067 D13 (``c58e36e64``):
    a date-time with a ``Z`` or an offset (``2024-02-03T04:05:06Z``, the form
    Word writes) is that instant; one with none (``2025-06-07T08:09:10``) and a
    date alone (``2024-02-03``, also ``2024`` and ``2024-02``, W3CDTF's coarser
    forms) are **UTC**, midnight for a date, as Word 365 reads a core
    property's (docx4j CR-067 section 12.2, its date-forms probe in Sydney:
    created ``2024-02-03`` showed 11:00 at UTC+11, modified
    ``2025-06-07T08:09:10`` 18:09 at UTC+10). Not the local clock, which is what
    a revision's ``w:date`` means instead (:func:`docx4j_py.model.values.date_time`).
    CR-025 Phase B2: until then these two forms were read nominally (D9, Q8),
    pending that probe. A date alone with an offset (``2024-02-03Z``) is
    midnight there, as XML Schema's ``xs:date`` reads it. Seconds are
    optional, a fraction of any length is cut to microseconds.
    """
    if value is None:
        return None
    found = _W3CDTF.fullmatch(value.strip())
    if found is None:
        return None
    year, month, day, hour, minute, second, fraction, offset = found.groups()
    try:
        when = datetime.datetime(  # noqa: DTZ001 - the zone is the offset's, else UTC, applied below
            int(year),
            int(month or 1),
            int(day or 1),
            int(hour or 0),
            int(minute or 0),
            int(second or 0),
            int((fraction or "0")[:6].ljust(6, "0")),
        )
    except ValueError:
        return None
    if offset is None or offset in ("Z", "z"):
        zone = datetime.UTC  # no offset: UTC, as Word reads a core property's (D13)
    else:
        sign = -1 if offset[0] == "-" else 1
        hours, minutes = int(offset[1:3]), int(offset[4:6])
        if hours > 14 or minutes > 59:
            return None
        zone = datetime.timezone(sign * datetime.timedelta(hours=hours, minutes=minutes))
    return when.replace(tzinfo=zone).astimezone(datetime.UTC)
