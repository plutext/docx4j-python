"""A main part's root, and a ``w:body`` given twice, read as Word reads them (CR-039 Phase B).

Word 365 opens a ``word/document.xml`` whose root isn't ``w:document`` more often
than a parse does, with no one rule for it. Measured through the Word bridge on
16.0.20326.20232 (CR-039 sections 2.2 and 2.4), and read here in this order:

* ``mc:AlternateContent``: the branch Word takes
  (:func:`~docx4j_py.openpackaging.mc_load_mode.branch_taken`). A ``w:document`` it
  holds becomes the root, and two are joined as two ``w:body`` are. Nothing selected
  is an empty document. What else it holds is read by the rows below, as it would
  be at the root, so a ``w:hdr`` in a taken ``mc:Choice`` is ``or-50``.
* The roots Word refuses, by name and by what they hold (WD REQ-318): ``w:footnotes``,
  ``w:endnotes``, Word 2003's ``w:wordDocument``, Flat OPC's ``pkg:package``, and an
  unknown ``w:`` root, or one in no namespace or one not understood, holding a
  ``w:document``. :class:`RootRefused`, so the part is refused as one that isn't well
  formed is (CR-031). They're listed and never "anything else", since ``w:madeUp``
  holding a ``w:p`` opens and what decides isn't known.
* A bare ``w:body``, wrapped in a ``w:document``.
* A ``w:`` name the model doesn't hold, wrapped in a ``w:document`` and a ``w:body``,
  where CR-037's pass takes the tags off. Measured on ``w:madeUp`` holding one ``w:p``.
* Any other root yields nothing (WD REQ-320): a ``w:document`` with no body, what the
  root held in ``part.skipped``, and :class:`UnexpectedRoot` to say so. A header's
  words in a footnote's separator aren't a reading to copy.

Only the main part is measured, so only it asks (``XmlPart.reads_root``).

**Two ``w:body``** (``rc-10``) join in order, one section: the later's. The fork keeps
the first and reports the second, so the skipped report triggers :func:`repair`, inside
:func:`~docx4j_py.openpackaging.read_through.repair_bytes`'s one repair and one reparse.

**The report** is core-ts CR-028 section 3.3's. A part read as Word reads it has lost
nothing, so it gets one :class:`ReadAsWord` with a count. What goes whole goes in
``part.skipped`` beside it, an earlier ``w:sectPr`` or a branch not taken, so
``strict=True`` refuses those and a part that only joined loads. Both codes are
proposals, not yet agreed with core-ts and docx4j (their D8 and D11).
"""

from __future__ import annotations

import dataclasses
from collections.abc import Iterable
from typing import Any, ClassVar

from docx4j_xsdata.formats.dataclass.parsers.skipped import SkippedNode
from lxml import etree

from docx4j_py.el_index import WML_ELEMENTS
from docx4j_py.openpackaging import breaks
from docx4j_py.openpackaging.mc_load_mode import branch_taken

__all__ = ["ReadAsWord", "RootRefused", "UnexpectedRoot", "read_root", "repair", "triggers"]

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _w(local: str) -> str:
    return f"{{{_W}}}{local}"


_DOCUMENT = _w("document")
_BODY = _w("body")
_SECT_PR = _w("sectPr")
_ALTERNATE_CONTENT = f"{{{_MC}}}AlternateContent"

#: Refused by name (``or-12``, ``or-13``, ``or-15``, ``rw-30``, ``rw-31``). A root holding a ``w:document`` is the other row.
_REFUSED = frozenset(
    {
        _w("footnotes"),
        _w("endnotes"),
        "{http://schemas.microsoft.com/office/word/2003/wordml}wordDocument",
        "{http://schemas.microsoft.com/office/2006/xmlPackage}package",
    }
)


class RootRefused(ValueError):
    """A root Word refuses to open. ``XmlPart`` turns it into the part's ``Docx4JException``."""


@dataclasses.dataclass(frozen=True, slots=True)
class ReadAsWord:
    """A part read as Word reads it, though it isn't as written: ``part.read_as_word``, one per part.

    Attributes:
        part_name: the part, ``/word/document.xml``.
        root: the root as written, where it wasn't ``w:document``.
        children: the containers joined into another.
    """

    code: ClassVar[str] = "part.read_as_word"

    part_name: str
    root: str | None
    children: int

    @property
    def message(self) -> str:
        """``/word/document.xml: root mc:AlternateContent, 1 children joined; read as Word reads them``."""
        said = [f"root {self.root}"] if self.root else []
        said += [f"{self.children} children joined"] if self.children else []
        return f"{self.part_name}: {', '.join(said)}; read as Word reads them"

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view, for the MCP server: ``code``, the three fields and ``message``."""
        return {
            "code": self.code,
            "part_name": self.part_name,
            "root": self.root,
            "children": self.children,
            "message": self.message,
        }


@dataclasses.dataclass(frozen=True, slots=True)
class UnexpectedRoot:
    """A part whose root isn't one it can read, left as it is: ``part.unexpected_root``, one per part.

    Attributes:
        part_name: the part, ``/word/document.xml``.
        root: the root, or what Word's branch of an ``mc:AlternateContent`` root held.
    """

    code: ClassVar[str] = "part.unexpected_root"

    part_name: str
    root: str

    @property
    def message(self) -> str:
        """``/word/document.xml: root w:hdr, not w:document; what it holds isn't read``."""
        return f"{self.part_name}: root {self.root}, not w:document; what it holds isn't read"

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view, for the MCP server: ``code``, the two fields and ``message``."""
        return {"code": self.code, "part_name": self.part_name, "root": self.root, "message": self.message}


def _name(tag: str) -> str:
    """`tag` as a reader knows it: ``w:hdr``, ``mc:AlternateContent``, ``foo``, else its Clark name."""
    if tag.startswith(f"{{{_W}}}"):
        return f"w:{etree.QName(tag).localname}"
    if tag.startswith(f"{{{_MC}}}"):
        return f"mc:{etree.QName(tag).localname}"
    return tag


def _elements(parent: etree._Element) -> list[etree._Element]:
    """`parent`'s child elements, without its comments and processing instructions."""
    return list(parent.iterchildren(etree.Element))


def _join_bodies(document: etree._Element, dropped: list[SkippedNode] | None) -> int:
    """Each later ``w:body`` of `document` joined to its first, one section the last stated: how many joined."""
    first, *later = document.findall(_BODY) or [None]
    for body in later:
        held = _elements(body)
        ended = first.find(_SECT_PR)
        if ended is not None:
            if any(child.tag == _SECT_PR for child in held):
                if dropped is not None:
                    dropped.append(breaks.skipped_node(ended))
            else:
                held.append(ended)
            first.remove(ended)
        first.extend(held)
        document.remove(body)
    return len(later)


def _join_documents(documents: list[etree._Element], dropped: list[SkippedNode]) -> tuple[etree._Element, int]:
    """`documents` as the first, every ``w:body`` joined: it, and how many containers joined."""
    first, *later = documents
    for document in later:
        for child in _elements(document):
            if child.tag == _BODY:
                first.append(child)
            else:
                dropped.append(breaks.skipped_node(child))
    return first, _join_bodies(first, dropped)


def _empty_document() -> etree._Element:
    """Word's reading of a root with nothing selected (``rw-24``)."""
    document = etree.Element(_DOCUMENT, nsmap={"w": _W})
    etree.SubElement(document, _BODY)
    return document


def _taken(root: etree._Element, dropped: list[SkippedNode]) -> tuple[etree._Element, int]:
    """What `root` comes to once an ``mc:AlternateContent`` is read: the element, and how many containers joined."""
    if root.tag != _ALTERNATE_CONTENT:
        return root, 0
    # REQ-305's first mc:Choice, Word's nothing selected: the WD's open ruling is this line
    branch = branch_taken(root)
    dropped.extend(breaks.skipped_node(child) for child in _elements(root) if child is not branch)
    held = [] if branch is None else _elements(branch)
    if not held:
        return _empty_document(), 0
    documents = [child for child in held if child.tag == _DOCUMENT]
    if documents:
        return _join_documents(documents, dropped)
    dropped.extend(breaks.skipped_node(child) for child in held[1:])
    return _taken(held[0], dropped)


def _wrapped(element: etree._Element) -> etree._Element:
    """A bare ``w:body`` in a ``w:document``, any other element in a ``w:body`` too."""
    document = etree.Element(_DOCUMENT, nsmap=element.nsmap)
    (document if element.tag == _BODY else etree.SubElement(document, _BODY)).append(element)
    return document


def _written(element: etree._Element) -> bytes:
    """`element` as a part's bytes, whatever it stood in, with the declarations it inherited."""
    element.tail = None
    return breaks.write(element)


def read_root(data: bytes, part_name: str) -> tuple[bytes, Any, list[SkippedNode]]:
    """A main part's bytes as Word reads its root, and what to say of that.

    Args:
        data: the part's bytes.
        part_name: the part, for the warning.

    Returns:
        ``(bytes, warning, dropped)``: `data` itself where the root is
        ``w:document``; a :class:`ReadAsWord`, an :class:`UnexpectedRoot` or None;
        and a ``SkippedNode`` for what went whole.

    Raises:
        RootRefused: a root Word refuses.
        etree.XMLSyntaxError: `data` isn't well formed.
    """
    root = breaks.parse(data)
    if root.tag == _DOCUMENT:
        return data, None, []
    dropped: list[SkippedNode] = []
    element, joined = _taken(root, dropped)
    written = _name(root.tag)
    if element.tag == _DOCUMENT:
        return _written(element), ReadAsWord(part_name, written, joined), dropped
    if _refused(element):
        raise RootRefused(f"its root is {_name(element.tag)}, which Word refuses to open")
    if element.tag == _BODY or _unknown(element.tag):
        return _written(_wrapped(element)), ReadAsWord(part_name, written, joined), dropped
    dropped.extend(breaks.skipped_node(child) for child in _elements(element))
    return breaks.write(etree.Element(_DOCUMENT, nsmap={"w": _W})), UnexpectedRoot(part_name, _name(element.tag)), dropped


def _unknown(tag: str) -> bool:
    """Whether `tag` is a ``w:`` name the model doesn't hold, CR-037's meaning."""
    return tag.startswith(f"{{{_W}}}") and etree.QName(tag).localname not in WML_ELEMENTS


def _refused(root: etree._Element) -> bool:
    """Whether Word refuses `root`, by name or by what it holds (WD REQ-318)."""
    from docx4j_py.child import understood_namespaces

    if root.tag in _REFUSED:
        return True
    namespace = etree.QName(root).namespace
    foreign = namespace is None or namespace not in understood_namespaces()
    return root.find(_DOCUMENT) is not None and (foreign or _unknown(root.tag))


def triggers(skipped: Iterable[Any]) -> bool:
    """Whether the fork's skipped report names a second ``w:body`` in a ``w:document``."""
    return any(
        getattr(item, "qname", None) == _BODY and (getattr(item, "parent_qname", None) or "").endswith("document")
        for item in skipped
    )


def repair(root: etree._Element, dropped: list[SkippedNode] | None = None) -> int:
    """Each ``w:body`` after the first in `root` joined to it, in place: how many joined."""
    return _join_bodies(root, dropped) if root.tag == _DOCUMENT else 0
