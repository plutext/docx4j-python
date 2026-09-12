"""``PartName``: an OPC part name. docx4j ``org.docx4j.openpackaging.parts.PartName``.

CR-002 section 5.1. A value object over a string that begins with ``/``, has no
trailing ``/``, no empty segments and no segment that ends in a dot; the
characters of a segment are RFC 3986 ``pchar`` or a percent escape. Equality and
hashing are **case-insensitive** (OPC [M1.12], and decided question 6), and the
spelling the package used is kept for writing, so a part Word called
``/word/Header1.xml`` is found by ``/word/header1.xml`` and still written as
``Header1.xml``.

Three static operations carry the whole relationship-resolution burden:

``resolve(source, target)``
    docx4j ``URIHelper.resolvePartUri``: a relationship's target read relative
    to the *directory* of the part that declared it, ``.`` and ``..``
    collapsed. ``TargetMode="External"`` targets never reach it.
``relativize(source, target)``
    docx4j ``URIHelper.relativizeURI``: the inverse, for writing a relationship.
``rels_for(name)`` / ``source_of_rels(name)``
    ``/word/document.xml`` <-> ``/word/_rels/document.xml.rels``, and ``/`` <->
    ``/_rels/.rels``.
"""

from __future__ import annotations

import re

from docx4j_py.openpackaging.exceptions import InvalidFormatException

__all__ = ["CONTENT_TYPES_NAME", "PACKAGE_RELS_NAME", "PartName"]

#: The content types stream. Not a part: it has no content type of its own and
#: is never in ``OpcPackage.parts``. docx4j ``ContentTypeManager.CONTENT_TYPES_PART_NAME``.
CONTENT_TYPES_NAME = "[Content_Types].xml"

#: The package relationships part, docx4j ``URIHelper.PACKAGE_RELATIONSHIPS_ROOT_URI``.
PACKAGE_RELS_NAME = "/_rels/.rels"

_SCHEME = re.compile(r"^[a-zA-Z][a-zA-Z0-9+.-]*:")
_RELS = re.compile(r"/_rels/[^/]*\.rels$")
_RELS_SPLIT = re.compile(r"^(.*)/_rels/([^/]*)\.rels$")
_BAD_CHAR = re.compile(r"[\\#?\s]")
_BAD_ESCAPE = re.compile(r"%(?![0-9A-Fa-f]{2})")
_SLASH_ESCAPE = re.compile(r"%2[fF]|%5[cC]")
_ALL_DOTS = re.compile(r"^\.+$")


class PartName:
    """An OPC part name. Immutable, hashable, case-insensitively equal."""

    __slots__ = ("_key", "_name")

    def __init__(self, name: str, *, _checked: bool = False) -> None:
        """Prefer :meth:`of`; this constructor validates unless told not to."""
        if not _checked:
            name = _validate(name)
        self._name = name
        self._key = name.lower()

    # -- construction ------------------------------------------------------

    @classmethod
    def of(cls, name: str | PartName) -> PartName:
        """Validate and wrap a part name; a missing leading ``/`` is added."""
        if isinstance(name, PartName):
            return name
        return cls(name)

    #: The package itself, the source of the package relationships. Not a legal
    #: part name; docx4j's ``URIHelper.PACKAGE_ROOT_PART_NAME``, built with
    #: ``checkConformance=false`` for exactly this reason.
    ROOT: PartName

    # -- the string --------------------------------------------------------

    @property
    def name(self) -> str:
        """The name with its leading ``/``, as the package spells it."""
        return self._name

    @property
    def key(self) -> str:
        """The lower-cased name; the key of every map in the engine."""
        return self._key

    @property
    def store_name(self) -> str:
        """The name as a container stores it: no leading ``/``."""
        return self._name[1:]

    @property
    def extension(self) -> str:
        """``'xml'`` for ``/word/document.xml``; ``''`` when there is no dot.

        docx4j uses the **last** dot here (``PartName.getExtension``) and the
        *first* one in ``ContentTypeManager.getPart``; the last is right and is
        what this uses throughout. The difference only shows on a part name
        with a dot in a directory segment, which Word does not write.
        """
        last = self._name[self._name.rfind("/") + 1 :]
        dot = last.rfind(".")
        return last[dot + 1 :] if dot > -1 else ""

    @property
    def directory(self) -> str:
        """``/word`` for ``/word/document.xml``; ``/`` at the top."""
        i = self._name.rfind("/")
        return "/" if i == 0 else self._name[:i]

    @property
    def file_name(self) -> str:
        """``document.xml`` for ``/word/document.xml``."""
        return self._name[self._name.rfind("/") + 1 :]

    @property
    def is_relationships_part(self) -> bool:
        """True for ``/_rels/.rels`` and ``/word/_rels/document.xml.rels``."""
        return _RELS.search(self._key) is not None

    # -- identity ----------------------------------------------------------

    def __eq__(self, other: object) -> bool:
        """Case-insensitive equality, OPC [M1.12]; a plain string works too."""
        if isinstance(other, PartName):
            return self._key == other._key
        if isinstance(other, str):
            return self._key == other.lower()
        return NotImplemented

    def __hash__(self) -> int:
        """Hash of the lower-cased name, so it agrees with :meth:`__eq__`."""
        return hash(self._key)

    def __lt__(self, other: PartName) -> bool:
        """Order by the lower-cased name, as docx4j's ``compareTo`` does."""
        return self._key < other._key

    def __str__(self) -> str:
        """The name itself."""
        return self._name

    def __repr__(self) -> str:
        """``PartName('/word/document.xml')``."""
        return f"PartName({self._name!r})"

    # -- relationships -----------------------------------------------------

    @staticmethod
    def rels_for(source: str | PartName) -> PartName:
        """The relationships part of a part; ``/_rels/.rels`` for the package.

        docx4j ``PartName.getRelationshipsPartName``.
        """
        text = source.name if isinstance(source, PartName) else source
        if text in ("/", ""):
            return PartName(PACKAGE_RELS_NAME, _checked=True)
        if not text.startswith("/"):
            text = "/" + text
        pos = text.rfind("/")
        return PartName(f"{text[:pos]}/_rels{text[pos:]}.rels", _checked=True)

    @staticmethod
    def source_of_rels(rels: str | PartName) -> PartName:
        """The part a relationships part belongs to; :attr:`ROOT` for the package."""
        text = rels.name if isinstance(rels, PartName) else rels
        match = _RELS_SPLIT.match(text)
        if match is None:
            raise InvalidFormatException(f"Not a relationships part name: {text}")
        head, tail = match.group(1), match.group(2)
        if head == "" and tail == "":
            return PartName.ROOT
        return PartName(f"{head}/{tail}")

    @staticmethod
    def resolve(source: str | PartName, target: str) -> PartName:
        """Resolve a relationship target against the part that declared it.

        docx4j ``URIHelper.resolvePartUri``. The target is read relative to the
        *directory* of `source` unless it begins with ``/``; ``.`` is dropped
        and ``..`` pops a segment. Only for internal targets: an
        ``External`` one is a URI and is never a part name.
        """
        text = source.name if isinstance(source, PartName) else source
        base = "/" if text == "/" else text[: text.rfind("/") + 1]
        path = target if target.startswith("/") else base + target
        out: list[str] = []
        for segment in path.split("/"):
            if segment in ("", "."):
                continue
            if segment == "..":
                if out:
                    out.pop()
            else:
                out.append(segment)
        return PartName("/" + "/".join(out))

    @staticmethod
    def relativize(source: str | PartName, target: str | PartName) -> str:
        """The relationship target to write for a part, seen from `source`.

        docx4j ``URIHelper.relativizeURI``, with its one wrinkle already
        applied: from the package root the answer is the name without its
        leading ``/`` (docx4j returns the absolute name and both of its callers
        strip the slash afterwards).
        """
        s = source.name if isinstance(source, PartName) else source
        t = target.name if isinstance(target, PartName) else target
        if s == "/":
            return t[1:]
        from_segments = [p for p in s[1 : s.rfind("/") + 1].split("/") if p]
        to_segments = t[1:].split("/")
        i = 0
        while i < len(from_segments) and i < len(to_segments) - 1:
            if from_segments[i] != to_segments[i]:
                break
            i += 1
        return "../" * (len(from_segments) - i) + "/".join(to_segments[i:])


def _validate(name: str) -> str:
    """Check a part name against OPC Annex A and return it with its slash."""
    if not isinstance(name, str) or name == "":
        raise InvalidFormatException("A part name shall not be empty [M1.1]")
    if _SCHEME.match(name):
        raise InvalidFormatException(f"Absolute URI forbidden: {name}")
    full = name if name.startswith("/") else "/" + name
    if full == "/":
        raise InvalidFormatException("A part name shall not be empty [M1.1]")
    if full.endswith("/"):
        raise InvalidFormatException(
            f"A part name shall not have a forward slash as the last character [M1.5]: {full}"
        )
    for segment in full[1:].split("/"):
        if segment == "":
            raise InvalidFormatException(
                f"A part name shall not have empty segments [M1.3]: {full}"
            )
        if segment.endswith("."):
            raise InvalidFormatException(
                f"A segment shall not end with a dot ('.') character [M1.9]: {full}"
            )
        if _ALL_DOTS.match(segment):
            raise InvalidFormatException(
                f"A segment shall include at least one non-dot character [M1.10]: {full}"
            )
        if _BAD_CHAR.search(segment):
            raise InvalidFormatException(
                "A segment shall not hold any characters other than pchar characters "
                f"[M1.6]: {full}"
            )
        if _BAD_ESCAPE.search(segment):
            raise InvalidFormatException(
                f"The segment {segment} contains an invalid encoded character"
            )
        if _SLASH_ESCAPE.search(segment):
            raise InvalidFormatException(
                "A segment shall not contain a percent-encoded forward or backward "
                f"slash [M1.7]: {full}"
            )
    return full


PartName.ROOT = PartName("/", _checked=True)
