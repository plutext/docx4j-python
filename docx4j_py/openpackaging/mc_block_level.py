"""``mc:AlternateContent`` at block, row and cell level, on load. docx4j ``McBlockLevel`` and ``McSplice``.

CR-012 Phase A (section 4.1 item 3), the port of docx4j CR-051 phase 1. Since
the schema refresh to docx4j ``2853b7280`` every block, row and cell container
of WordprocessingML admits ``mc:AlternateContent``: the body (a glossary
entry's too), a header or a footer, a cell, a note, a comment, a text box, a
block-level content control or custom XML element (block containers); a table
and a row-level control or custom XML element (row containers); a row and a
cell-level control or custom XML element (cell containers). What a load does
with such an element is the **mode**:

``"preserve"`` (the default here, CR-012 Q7)
    nothing is done: the element is kept whole, every branch, and a reader
    selects a branch at use, as at run and paragraph level since CR-002. A
    part saved unchanged is canonically identical to its source.
``"word"``
    each such element is replaced, before the part is parsed, by the content
    of the branch Word reads: the first ``mc:Choice`` whose ``Requires``
    namespaces are all in docx4j's published understood list (matched by URI,
    through the declarations in scope at the Choice; WD REQ-302 to REQ-304),
    else the ``mc:Fallback``. The tree then holds what Word's holds, and a save
    writes what Word's save writes (CR-051 phase 0), text boxes apart (its
    section 3.5).

In ``word`` mode, as docx4j's splice does:

* the elements are taken **outermost first**; one inside a branch that was
  discarded, or inside an element kept whole, stays whole in there;
* an element with no understood Choice and no Fallback is **kept whole**
  (WD REQ-305), and one with no branch at all is **removed**;
* the selected Choice's ``Requires`` prefixes are **added to the root's
  ``mc:Ignorable``** where it does not list them already, as the conventional
  prefix of each namespace (the writer declares it from the source or the
  prefix table);
* each element taken out is **recorded** (:class:`ResolvedAlternateContent`),
  whole and with where it stood, so that nothing is lost to inspection.

Run- and paragraph-level elements are not governed: they are kept whole in
both modes. The splice runs on an lxml tree of the part's bytes before the
fork's parser sees them, as :mod:`docx4j_py.openpackaging.mce`'s preprocessor
does, so the selected content is parsed typed by the ordinary path; a part
whose bytes do not hold ``AlternateContent`` pays one substring test.

Departures from docx4j ``9db1a3cbd`` (its CR-051 phase 2), recorded in CR-012
sections 9 and 10: the default is ``preserve`` (CR-012 Q7, as docx4j's D5
decides for 17.4.0; its 17.3.x code defaults to ``word``); the splice is on the
source's lxml tree rather than on the unmarshalled objects, so a record's
element is a parse of its own, not the objects the tree held. ``preserve``
keeps row- and cell-level elements whole, as docx4j's phase 2 does (its phase 1
still resolved them in that mode).
"""

from __future__ import annotations

import dataclasses
from typing import Any

from lxml import etree

__all__ = [
    "MODES",
    "PRESERVE",
    "WORD",
    "ResolvedAlternateContent",
    "check_mode",
    "splice_word",
]

#: Each element kept whole (the default, CR-012 Q7).
PRESERVE = "preserve"
#: Each element resolved, on load, to the branch Word reads.
WORD = "word"
#: The two modes.
MODES = (PRESERVE, WORD)

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_AC = f"{{{_MC}}}AlternateContent"
_CHOICE = f"{{{_MC}}}Choice"
_FALLBACK = f"{{{_MC}}}Fallback"
_MC_WRAPPERS = frozenset({_AC, _CHOICE, _FALLBACK})
_IGNORABLE = f"{{{_MC}}}Ignorable"

#: The block containers: each holds blocks whatever stands around it.
_BLOCK = frozenset(
    f"{{{_W}}}{local}"
    for local in (
        "body",
        "hdr",
        "ftr",
        "tc",
        "footnote",
        "endnote",
        "comment",
        "txbxContent",
        "docPartBody",
    )
)
_TBL = f"{{{_W}}}tbl"
_TR = f"{{{_W}}}tr"
_SDT = f"{{{_W}}}sdt"
_SDT_CONTENT = f"{{{_W}}}sdtContent"
_CUSTOM_XML = f"{{{_W}}}customXml"

#: The levels a container's content is at, and what a ``w:sdt`` straight in a branch there is.
BLOCK, ROW, CELL = "block", "row", "cell"

#: Property elements, which stand in a container's children but not in its content list: a
#: record's index and address count the content, as the model's lists hold it.
_PROPERTIES = frozenset(
    f"{{{_W}}}{local}"
    for local in (
        "pPr",
        "rPr",
        "tblPr",
        "tblGrid",
        "tblPrEx",
        "trPr",
        "tcPr",
        "sdtPr",
        "sdtEndPr",
        "customXmlPr",
        "smartTagPr",
        "sectPr",
    )
)

#: The substring a part must hold for the splice to look at it.
_MARKER = b"AlternateContent"

_PARSER = etree.XMLParser(
    remove_blank_text=False, resolve_entities=False, huge_tree=True
)


def check_mode(value: Any) -> str:
    """`value` if it names a mode; ``ValueError`` otherwise."""
    if value not in MODES:
        raise ValueError(f"mc_block_level must be one of {MODES!r}, not {value!r}")
    return value


@dataclasses.dataclass(frozen=True, slots=True)
class ResolvedAlternateContent:
    """An ``mc:AlternateContent`` the ``word`` mode took out of a part. docx4j ``ResolvedAlternateContent``.

    Where it stood is an address worked out when it was taken out, not a live
    reference, so the record outlives an edit to the tree.

    Attributes:
        element: the element taken out, whole: every ``mc:Choice`` and the
            ``mc:Fallback``, as a typed ``docx4j_py.mce.AlternateContent``
            parsed on its own (the objects are not the tree's: the tree's
            promoted content was parsed from the same bytes, in place).
        branch: the branch of `element` whose content took its place (an
            ``AlternateContentChoice`` or the ``AlternateContentFallback``);
            None where the element had no branch at all and was removed.
        container_name: the container's element name, prefixed
            (``"w:body"``, ``"w:tc"``, ``"w:sdtContent"``, ``"w:tbl"``).
        level: what the container holds: ``"block"``, ``"row"`` or ``"cell"``.
        address: the container's address from the part's root, each step an
            element name with its index among its parent's content, property
            elements left out (``"w:document/w:body[0]/w:tbl[3]/w:tr[0]/w:tc[1]"``).
        index: the index in the container's content at which the element
            stood, and the promoted content now starts.
        promoted: how many elements the branch put in its place (0 where it
            was removed).
    """

    element: Any
    branch: Any
    container_name: str
    level: str
    address: str
    index: int
    promoted: int

    def __repr__(self) -> str:
        """``mc:AlternateContent at w:document/w:body[0] index 2: Choice Requires="w16sdtfl", 1 promoted``."""
        branch = self.branch
        if branch is None:
            which = "none (removed)"
        elif type(branch).__name__ == "AlternateContentChoice":
            which = f'Choice Requires="{branch.requires}"'
        else:
            which = "Fallback"
        return f"mc:AlternateContent at {self.address} index {self.index}: {which}, {self.promoted} promoted"


@dataclasses.dataclass(slots=True)
class _Taken:
    """What the splice took out, before the record's element is parsed."""

    xml: bytes
    branch: (
        int | None
    )  # the branch's position among the element's branches, choices then the fallback
    container_name: str
    level: str
    address: str
    index: int
    promoted: int


def splice_word(data: bytes) -> tuple[bytes, list[ResolvedAlternateContent]]:
    """A part's bytes with every block-, row- and cell-level element replaced by the branch Word reads.

    Returns:
        The bytes to parse (`data` itself where nothing was taken out) and a
        record of each element taken out, in the order taken: outermost first,
        in document order.
    """
    if _MARKER not in data:
        return data, []
    root = etree.fromstring(data, _PARSER)
    taken, prefixes = _splice(root)
    if not taken:
        return data, []
    _add_ignorable(root, prefixes)
    spliced = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    return spliced, [_record(item) for item in taken]


# ---------------------------------------------------------------------------
# the splice
# ---------------------------------------------------------------------------


def _splice(root: etree._Element) -> tuple[list[_Taken], list[str]]:
    from docx4j_py.child import understood_namespaces

    understood = understood_namespaces()
    taken: list[_Taken] = []
    prefixes: list[str] = []
    kept: set[etree._Element] = set()
    # document order is outermost first: a nested element is met after the one holding it
    for ac in list(root.iter(_AC)):
        if not _attached(ac, root, kept):
            continue  # in a discarded branch, or in an element kept whole: it stays whole in there
        container = ac.getparent()
        if container is None or container.tag in _MC_WRAPPERS:
            continue  # straight in a branch of an element this splice does not govern
        level = _level(container)
        if level is None:
            continue  # run or paragraph level, or a property: kept whole in both modes
        choices = [child for child in ac if child.tag == _CHOICE]
        fallback = next((child for child in ac if child.tag == _FALLBACK), None)
        name = _prefixed(container.tag)
        address = _address(container)
        index = _content_index(container, ac)
        if not choices and fallback is None:
            # no branch at all: nothing shows in any reader, and Word's save leaves it out
            taken.append(
                _Taken(etree.tostring(ac), None, name, level, address, index, 0)
            )
            _remove(ac)
            continue
        selected = next(
            (choice for choice in choices if _understood(choice, understood)), None
        )
        if selected is None and fallback is None:
            kept.add(ac)  # REQ-305: no understood Choice and no Fallback, kept whole
            continue
        branch = selected if selected is not None else fallback
        position = choices.index(selected) if selected is not None else len(choices)
        xml = etree.tostring(ac)
        promoted = _replace(ac, branch)
        taken.append(_Taken(xml, position, name, level, address, index, promoted))
        if selected is not None:
            for prefix in _requires_prefixes(selected):
                if prefix not in prefixes:
                    prefixes.append(prefix)
    return taken, prefixes


def _attached(
    ac: etree._Element, root: etree._Element, kept: set[etree._Element]
) -> bool:
    """Whether `ac` is still in the tree under `root`, and not inside an element kept whole."""
    node = ac.getparent()
    top = ac
    while node is not None:
        if node in kept:
            return False
        top = node
        node = node.getparent()
    return top is root


def _effective_container(element: etree._Element) -> etree._Element | None:
    """The first ancestor that is not ``mc:AlternateContent``, ``mc:Choice`` or ``mc:Fallback``."""
    node = element.getparent()
    while node is not None and node.tag in _MC_WRAPPERS:
        node = node.getparent()
    return node


def _level(container: etree._Element) -> str | None:
    """What a container's content is: blocks, rows or cells; None for anything else (a paragraph, a run, a property)."""
    tag = container.tag
    if tag in _BLOCK:
        return BLOCK
    if tag == _TBL:
        return ROW
    if tag == _TR:
        return CELL
    if tag == _SDT_CONTENT:
        sdt = container.getparent()
        if sdt is None or sdt.tag != _SDT:
            return None
        return _holder_level(sdt)
    if tag == _CUSTOM_XML:
        return _holder_level(container)
    return None


def _holder_level(holder: etree._Element) -> str | None:
    """The level of what a ``w:sdt`` or ``w:customXml`` holds: the level of the content it stands in."""
    container = _effective_container(holder)
    return _level(container) if container is not None else None


def _understood(choice: etree._Element, understood: frozenset[str]) -> bool:
    """WD REQ-302 to REQ-304: a ``Requires`` naming at least one prefix, each bound in scope to an understood namespace."""
    requires = (choice.get("Requires") or "").split()
    if not requires:
        return False
    in_scope = choice.nsmap
    return all(in_scope.get(prefix) in understood for prefix in requires)


def _requires_prefixes(choice: etree._Element) -> list[str]:
    """The prefix the root's ``mc:Ignorable`` names for each namespace a Choice requires: the conventional one, else its own."""
    from docx4j_py.namespaces import NAMESPACE_TO_PREFIX

    out: list[str] = []
    in_scope = choice.nsmap
    for prefix in (choice.get("Requires") or "").split():
        uri = in_scope.get(prefix)
        conventional = NAMESPACE_TO_PREFIX.get(uri) if uri is not None else None
        out.append(conventional or prefix)
    return out


def _replace(ac: etree._Element, branch: etree._Element) -> int:
    """`ac` replaced in its parent by `branch`'s children; how many elements that put there."""
    parent = ac.getparent()
    index = parent.index(ac)
    tail = ac.tail
    children = list(branch)
    for offset, child in enumerate(children):
        parent.insert(index + offset, child)
    if children:
        last = children[-1]
        if tail:
            last.tail = (last.tail or "") + tail
    elif branch.text and branch.text.strip():
        _append_text(parent, index, branch.text)
    parent.remove(ac)
    return sum(1 for child in children if isinstance(child.tag, str))


def _remove(ac: etree._Element) -> None:
    """`ac` taken out of its parent, its tail kept."""
    parent = ac.getparent()
    index = parent.index(ac)
    if ac.tail:
        _append_text(parent, index, ac.tail)
    parent.remove(ac)


def _append_text(parent: etree._Element, index: int, text: str) -> None:
    if index == 0:
        parent.text = (parent.text or "") + text
    else:
        previous = parent[index - 1]
        previous.tail = (previous.tail or "") + text


def _add_ignorable(root: etree._Element, prefixes: list[str]) -> None:
    """Each of `prefixes` the root's ``mc:Ignorable`` does not name, added; set where the root had none."""
    if not prefixes:
        return
    current = (root.get(_IGNORABLE) or "").split()
    missing = [prefix for prefix in prefixes if prefix not in current]
    if missing:
        root.set(_IGNORABLE, " ".join(current + missing))


# ---------------------------------------------------------------------------
# addresses
# ---------------------------------------------------------------------------


def _prefixed(tag: str) -> str:
    """``{uri}local`` with docx4j's conventional prefix for the URI, or the local name alone."""
    from docx4j_py.namespaces import NAMESPACE_TO_PREFIX

    uri, _, local = tag[1:].partition("}") if tag.startswith("{") else ("", "", tag)
    prefix = NAMESPACE_TO_PREFIX.get(uri) if uri else None
    return f"{prefix}:{local}" if prefix else local


def _content_index(parent: etree._Element, child: etree._Element) -> int:
    """`child`'s index among `parent`'s elements, its property elements left out."""
    index = 0
    for node in parent:
        if node is child:
            return index
        if isinstance(node.tag, str) and node.tag not in _PROPERTIES:
            index += 1
    return -1


def _address(element: etree._Element) -> str:
    """The element's address from the root: each step its name, and its content index where it has a parent."""
    steps: list[str] = []
    node: etree._Element | None = element
    while node is not None:
        parent = node.getparent()
        name = _prefixed(node.tag)
        steps.append(
            f"{name}[{_content_index(parent, node)}]" if parent is not None else name
        )
        node = parent
    return "/".join(reversed(steps))


# ---------------------------------------------------------------------------
# the record's element
# ---------------------------------------------------------------------------


def _record(item: _Taken) -> ResolvedAlternateContent:
    """The record, its element parsed whole, a ``w:sdt`` straight in a branch typed by the container's level."""
    from docx4j_py.mce import AlternateContent
    from docx4j_py.runtime import parser

    element = parser(
        lenient=True, skipped_report=False, wildcard_class=_record_hook(item.level)
    ).from_bytes(item.xml, AlternateContent)
    branch = None
    if item.branch is not None:
        branches = [
            *element.choice,
            *([element.fallback] if element.fallback is not None else []),
        ]
        branch = branches[item.branch]
    return ResolvedAlternateContent(
        element,
        branch,
        item.container_name,
        item.level,
        item.address,
        item.index,
        item.promoted,
    )


def _record_hook(level: str) -> Any:
    """The fork's ``wildcard_class`` for a record's own parse: no container above it, so the level says.

    A ``w:sdt``, ``w:tr`` or ``w:tc`` straight in one of the element's branches
    is bound as a container of the record's level binds it: a body for block
    level, a table for row level, a row for cell level (so a cell-level
    branch's ``w:tc`` is a ``Tc``, CR-012 Phase A2).
    """
    from docx4j_py import wml
    from docx4j_py.runtime import branch_child_class

    container = {ROW: wml.Tbl, CELL: wml.Tr}.get(level, wml.Body1)

    def hook(qname: str, ancestors: tuple[type, ...]) -> type | None:
        found = branch_child_class(qname, ancestors)
        if found is not None:
            return found
        if ancestors and all(cls.__module__ == "docx4j_py.mce" for cls in ancestors):
            return branch_child_class(qname, (*ancestors, container))
        return None

    return hook
