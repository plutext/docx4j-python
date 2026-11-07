"""A main part's root, and a child given more than once, read as Word reads them (CR-039).

Word 365 opens a ``word/document.xml`` whose root isn't ``w:document`` more often than
a parse does, with no one rule for it. Measured through the Word bridge on
16.0.20326.20232 (CR-039 sections 2 and 8.5), and read here as WD REQ-317 to REQ-320 say:

* ``mc:AlternateContent``: the branch Word takes
  (:func:`~docx4j_py.openpackaging.mc_load_mode.branch_taken`). A ``w:document`` it
  holds becomes the root, and two are joined as two ``w:body`` are. Nothing selected
  is an empty document. What else it holds is read by the rows below, as it would be
  at the root.
* The roots Word refuses (REQ-318): ``w:footnotes``, ``w:endnotes``, ``w:tc``, Word 2003's
  ``wordDocument``, Flat OPC's ``pkg:package``; an unknown ``w:`` root holding a
  ``w:document`` or, directly, a ``w:r``; a root in no namespace, or one neither
  understood nor in its ``mc:Ignorable``, holding a ``w:document`` or a ``w:p``.
  :class:`RootRefused`, so the part is refused as one that isn't well formed is (CR-031).
* The roots Word reads (REQ-319): a ``w:body``, a ``w:p``, a ``w:sdtContent``, and an
  unknown ``w:`` root holding ``w:p``, ``w:tbl``, a block ``w:sdt`` or a further unknown
  element, wrapped so that CR-037's pass takes its tags off. A ``w:sectPr`` directly in
  it has no effect, a ``w:body`` in it is the body, and empty it is an empty document.
* Any other root yields nothing (REQ-320): a ``w:document`` with no body, what the root
  held in ``part.skipped``, and :class:`UnexpectedRoot` to say so. A header's words in
  a footnote's separator aren't a reading to copy.

Only the main part is measured, so only it asks (``XmlPart.reads_root``).

**A child given more than once** (REQ-322). The fork keeps the first and reports the
rest, so the skipped report triggers :func:`repair`, inside
:func:`~docx4j_py.openpackaging.read_through.repair_bytes`'s one repair and one reparse.
``w:body``, and ``w:sdtContent`` in a block ``w:sdt``, are joined in order, one section
the last stated; a ``w:body`` in a ``w:body`` gives way to what it holds. ``w:pPr`` and
``w:rPr`` merge, the later value winning. Of ``w:jc``, ``w:b``, ``w:sz`` and ``w:tblPr`` the
last stands. Nothing else is measured, so a second ``w:txbxContent`` or ``w:tcPr`` stays
as it was, the first kept and the rest listed.

**The report** is core-ts CR-028 section 3.3's. A part read as Word reads it has lost
nothing, so it gets one :class:`ReadAsWord` with a count. What goes whole goes in
``part.skipped`` beside it, a value replaced, an earlier ``w:sectPr`` or a branch not
taken, so ``strict=True`` refuses those and a part that only joined loads.
``part.read_as_word`` is a proposal for core-ts's D8.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Iterable
from typing import Any, ClassVar

from docx4j_xsdata.formats.dataclass.parsers.skipped import SkippedNode
from lxml import etree

from docx4j_py.el_index import WML_ELEMENTS
from docx4j_py.openpackaging import breaks
from docx4j_py.openpackaging.mc_load_mode import branch_taken, required_namespaces

__all__ = ["ReadAsWord", "RootRefused", "UnexpectedRoot", "read_root", "repair", "triggers", "with_root"]

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _w(local: str) -> str:
    return f"{{{_W}}}{local}"


_DOCUMENT = _w("document")
_BODY = _w("body")
_SECT_PR = _w("sectPr")
_P = _w("p")
_R = _w("r")
_TBL = _w("tbl")
_TC = _w("tc")
_SDT = _w("sdt")
_SDT_CONTENT = _w("sdtContent")
_PPR = _w("pPr")
_RPR = _w("rPr")
_TBL_PR = _w("tblPr")
_ALTERNATE_CONTENT = f"{{{_MC}}}AlternateContent"
_IGNORABLE = f"{{{_MC}}}Ignorable"

#: What an unknown ``w:`` root reads through to, with a further unknown element (REQ-319).
_READ_IN_AN_UNKNOWN_ROOT = frozenset({_P, _TBL, _SDT})
#: Where a ``w:sdt`` is a block one, as far as joining its content goes.
_BLOCK_HOSTS = frozenset({_BODY, _TC})
#: The children REQ-322 reads, the ones the skipped report triggers on.
_MEASURED = frozenset(_w(local) for local in ("body", "sdtContent", "pPr", "rPr", "jc", "b", "sz", "sectPr", "tblPr"))
#: Of these, by the properties that hold them, the last stands.
_LAST_KEPT = ((_PPR, (_w("jc"),)), (_RPR, (_w("b"), _w("sz"))), (_TBL, (_TBL_PR,)))

#: Refused by name (``or-12``, ``or-13``, ``or-15``, ``mr-31``, ``rw-30``, ``rw-31``). What a root holds is the other rows.
_REFUSED = frozenset(
    {
        _w("footnotes"),
        _w("endnotes"),
        _TC,
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
        children: the repeated children read as one.
        nothing_selected: the namespaces an ``mc:AlternateContent`` root required, as URIs,
            where it selected nothing (REQ-305), which the text model raises as
            ``mce.nothing_selected``.
    """

    code: ClassVar[str] = "part.read_as_word"

    part_name: str
    root: str | None
    children: int
    nothing_selected: str | None = None

    @property
    def message(self) -> str:
        """``/word/document.xml: root mc:AlternateContent, 1 repeated children; read as Word reads them``."""
        said = [f"root {self.root}"] if self.root else []
        said += [f"{self.children} repeated children"] if self.children else []
        return f"{self.part_name}: {', '.join(said)}; read as Word reads them"

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view, for the MCP server: ``code``, the three fields and ``message``."""
        return {
            "code": self.code,
            "part_name": self.part_name,
            "root": self.root,
            "children": self.children,
            "nothing_selected": self.nothing_selected,
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


def _drop(parent: etree._Element, element: etree._Element, dropped: list[SkippedNode] | None) -> None:
    """`element` taken out of `parent`, and listed: it went whole."""
    if dropped is not None:
        dropped.append(breaks.skipped_node(element))
    parent.remove(element)


def _keep_last(parent: etree._Element, tag: str, dropped: list[SkippedNode] | None) -> int:
    """Only the last `tag` left in `parent`: how many went."""
    *earlier, _last = parent.findall(tag) or [None]
    for element in earlier:
        _drop(parent, element, dropped)
    return len(earlier)


def _flatten(body: etree._Element) -> int:
    """Each ``w:body`` in `body`, at any depth, given way to what it holds, in place: how many (``nd-20`` to ``nd-24``)."""
    flattened = 0
    while (inner := body.find(_BODY)) is not None:
        index = body.index(inner)
        body[index : index + 1] = _elements(inner)
        flattened += 1
    return flattened


def _join_bodies(document: etree._Element, dropped: list[SkippedNode] | None) -> int:
    """Each ``w:body`` of `document` joined to its first, one section the last stated, at the end: how many read as one."""
    bodies = document.findall(_BODY)
    if not bodies:
        return 0
    first, *later = bodies
    for body in later:
        first.extend(_elements(body))
        document.remove(body)
    read_as_one = len(later) + _flatten(first) + _keep_last(first, _SECT_PR, dropped)
    section = first.find(_SECT_PR)
    if section is not None:
        first.append(section)
    return read_as_one


def _join_contents(sdt: etree._Element) -> int:
    """Each ``w:sdtContent`` of `sdt` joined to its first: how many joined."""
    first, *later = sdt.findall(_SDT_CONTENT) or [None]
    for content in later:
        first.extend(_elements(content))
        sdt.remove(content)
    return len(later)


def _merge(parent: etree._Element, tag: str, dropped: list[SkippedNode] | None) -> int:
    """Each `tag` of `parent` merged into its first, the later value winning: how many merged."""
    first, *later = parent.findall(tag) or [None]
    for properties in later:
        for child in _elements(properties):
            for replaced in first.findall(child.tag):
                _drop(first, replaced, dropped)
            first.append(child)
        parent.remove(properties)
    return len(later)


def _join_documents(documents: list[etree._Element], dropped: list[SkippedNode]) -> tuple[etree._Element, int]:
    """`documents` as the first, every ``w:body`` joined: it, and how many read as one."""
    first, *later = documents
    for document in later:
        for child in _elements(document):
            if child.tag == _BODY:
                first.append(child)
            else:
                dropped.append(breaks.skipped_node(child))
    return first, _join_bodies(first, dropped)


def _empty_document() -> etree._Element:
    """Word's reading of a root with nothing selected (``rw-24``) and of an empty unknown root (``mr-16``)."""
    document = etree.Element(_DOCUMENT, nsmap={"w": _W})
    etree.SubElement(document, _BODY)
    return document


def _taken(root: etree._Element, dropped: list[SkippedNode], required: list[str]) -> tuple[etree._Element, int]:
    """What `root` comes to once an ``mc:AlternateContent`` is read: the element, and how many read as one.

    `required` gets the namespaces of one that selects nothing (REQ-305).
    """
    if root.tag != _ALTERNATE_CONTENT:
        return root, 0
    branch = branch_taken(root)
    dropped.extend(breaks.skipped_node(child) for child in _elements(root) if child is not branch)
    if branch is None and _elements(root):
        required.append(required_namespaces(root))
    held = [] if branch is None else _elements(branch)
    if not held:
        return _empty_document(), 0
    documents = [child for child in held if child.tag == _DOCUMENT]
    if documents:
        return _join_documents(documents, dropped)
    dropped.extend(breaks.skipped_node(child) for child in held[1:])
    return _taken(held[0], dropped, required)


def _wrapped(element: etree._Element) -> etree._Element:
    """A bare ``w:body`` in a ``w:document``, any other element in a ``w:body`` too."""
    document = etree.Element(_DOCUMENT, nsmap=element.nsmap)
    (document if element.tag == _BODY else etree.SubElement(document, _BODY)).append(element)
    return document


def _read_through(element: etree._Element, dropped: list[SkippedNode]) -> tuple[etree._Element | None, int]:
    """`element` as the ``w:document`` Word reads it as (REQ-319), or None, and how many read as one."""
    if element.tag in (_BODY, _P):
        return _wrapped(element), 0
    if element.tag == _SDT_CONTENT:
        element.tag = _BODY
        return _wrapped(element), 0
    if not _unknown(element.tag):
        return None, 0
    held = _elements(element)
    if not held:
        return _empty_document(), 0
    if any(child.tag == _BODY for child in held):
        document = etree.Element(_DOCUMENT, nsmap=element.nsmap)
        for child in held:
            if child.tag == _BODY:
                document.append(child)
            else:
                dropped.append(breaks.skipped_node(child))
        return document, _join_bodies(document, dropped)
    if not any(child.tag in _READ_IN_AN_UNKNOWN_ROOT or _unknown(child.tag) for child in held):
        return None, 0
    for section in element.findall(_SECT_PR):
        _drop(element, section, dropped)
    return _wrapped(element), 0


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
    written = _name(root.tag)
    required: list[str] = []
    element, read_as_one = _taken(root, dropped, required)
    if element.tag != _DOCUMENT:
        if _refused(element):
            raise RootRefused(f"its root is {_name(element.tag)}, which Word refuses to open")
        document, more = _read_through(element, dropped)
        if document is None:
            dropped.extend(breaks.skipped_node(child) for child in _elements(element))
            empty = etree.Element(_DOCUMENT, nsmap={"w": _W})
            return breaks.write(empty), UnexpectedRoot(part_name, _name(element.tag)), dropped
        element, read_as_one = document, read_as_one + more
    return _written(element), ReadAsWord(part_name, written, read_as_one, required[0] if required else None), dropped


def _unknown(tag: Any) -> bool:
    """Whether `tag` is a ``w:`` name the model doesn't hold, CR-037's meaning."""
    return isinstance(tag, str) and tag.startswith(f"{{{_W}}}") and etree.QName(tag).localname not in WML_ELEMENTS


def _refused(root: etree._Element) -> bool:
    """Whether Word refuses `root`, by name or by what it holds (WD REQ-318)."""
    from docx4j_py.child import understood_namespaces

    if root.tag in _REFUSED:
        return True
    holds_a_document = root.find(_DOCUMENT) is not None
    if _unknown(root.tag):
        return holds_a_document or root.find(_R) is not None
    namespace = etree.QName(root).namespace
    ignorable = {root.nsmap.get(prefix) for prefix in root.get(_IGNORABLE, "").split()}
    foreign = namespace is None or (namespace not in understood_namespaces() and namespace not in ignorable)
    return foreign and (holds_a_document or root.find(_P) is not None)


def triggers(skipped: Iterable[Any]) -> bool:
    """Whether the fork's skipped report names a child REQ-322 reads. A part that holds one only once changes nothing."""
    return any(getattr(item, "qname", None) in _MEASURED for item in skipped)


def repair(root: etree._Element, dropped: list[SkippedNode] | None = None) -> int:
    """Each child of a main part's `root` given more than once read as one, in place (REQ-322): how many read as one."""
    if root.tag != _DOCUMENT:
        return 0
    read_as_one = _join_bodies(root, dropped)
    for sdt in list(root.iter(_SDT)):
        if sdt.getparent().tag in _BLOCK_HOSTS:
            read_as_one += _join_contents(sdt)
    for host, tag in ((_P, _PPR), (_R, _RPR)):
        for parent in list(root.iter(host)):
            read_as_one += _merge(parent, tag, dropped)
    for host, tags in _LAST_KEPT:
        for parent in list(root.iter(host)):
            for tag in tags:
                read_as_one += _keep_last(parent, tag, dropped)
    return read_as_one


def with_root(root: Any, warnings: list[Any]) -> list[Any]:
    """`warnings` with the root's first, one :class:`ReadAsWord` a part however many passes made one."""
    if not isinstance(root, ReadAsWord):
        return [root, *warnings]
    read_as_one = sum(warning.children for warning in warnings if isinstance(warning, ReadAsWord))
    rest = [warning for warning in warnings if not isinstance(warning, ReadAsWord)]
    return [dataclasses.replace(root, children=root.children + read_as_one), *rest]
