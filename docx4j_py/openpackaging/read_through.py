"""An unknown ``w:`` element, read as Word 365 reads it. core-ts's CR-026 D7 and D8.

CR-037, on core-ts's Word check 76 (its CR-026 section 4.4), the WD's REQ-308 to
REQ-312. ``w:madeUp`` isn't a WordprocessingML element, and Word 365 doesn't treat
it as one thing. In a content host it takes the tags away and reads what they
held, at any depth, with no message. The model has no place for one, so the
fork's parse skips it and all it holds: every reader and every save went without
the words Word shows, with the loss in ``part.skipped``.

The rules, each decided on the tree as read and then applied:

1. an element in the ``w:`` namespace whose name the model doesn't hold
   (:data:`~docx4j_py.el_index.WML_ELEMENTS`), standing in one of the twenty hosts
   measured, or in such an element standing in one, gives way to what it held, in
   order (REQ-308; ``76t`` to ``76t33``). A tracked host's run stays tracked. Text
   it holds directly is absent (REQ-314);
2. so too directly in the ``mc:Choice`` or ``mc:Fallback`` Word takes, where the
   ``mc:AlternateContent`` stands in a ``w:p`` (REQ-313; ``76t44``, ``76t45``);
3. a ``w:r`` such an element holds directly goes to the head of its paragraph,
   straight after the ``w:pPr``, where the element stands in a paragraph's ``w:pPr``,
   in the ``w:pPr`` under its ``w:pPrChange``, or in a ``w:sectPr`` in its ``w:pPr``
   (REQ-311; ``76t27``, ``76t37`` to ``76t40``, ``76t50``). A ``w:sectPr`` that held
   one goes, with its section break, reported. Elsewhere what such an element holds
   stays in it, so the fork drops and reports that, as for a property (``76t5``);
4. a ``w:p`` such an element holds, where the element comes last in a ``w:p`` and a
   run comes before it, ends the host: its runs are the host's, and the host takes
   its properties if it has any, else keeps its own. A further ``w:p`` the element
   holds is a paragraph of its own, and an empty paragraph follows, with the host's
   properties (REQ-316; ``76t43``, ``76t56``, ``76t58``, ``76t68`` to ``76t70``). Word
   refuses every other such shape in a ``w:p``, so those stay the second parse's to
   skip and list.

Everything else stays as the parse leaves it: a property in ``w:pPr`` or ``w:rPr``
and anything in a ``w:t``, measured; a run in a ``w:rPr``, ``w:trPr``, ``w:tcPr``,
``w:tblPr`` or a block ``w:sdtPr``, files Word refuses (``76t28``, ``76t46`` to
``76t49``); and every host not measured. Adding a host later is one line of
:data:`_HOSTS`. An element of another namespace stops the walk upward, since Word
drops one with all it holds (``76k``), so a ``w:madeUp`` inside it isn't read
through. A ``w:`` element with a name the model does hold is never taken, though it's
in the wrong place: the skipped report can't tell the two apart, and that's what
the table is for.

Otherwise the pass doesn't look at what a host admits: what's left over is the
second parse's to skip and list, as any misplaced element is.

Each element taken is one :class:`SkippedNode`, so ``part.skipped`` lists it and
``strict=True`` refuses, though no words go: the tags are lost whatever the words
do, and core-ts reports the same (CR-037 Q1).

:func:`repair_bytes` is the one repair of a part: this pass and then CR-030's, on
one tree, then one parse, so a ``w:br`` such an element held directly is rule 1's of
CR-030 to place, as Word reads it (REQ-315; ``76t42``). :func:`triggers` reads the
fork's report. The fork doesn't report an element in a kept branch, since a wildcard
holds it, so :func:`in_a_branch` reads the bytes for one, at the cost of a substring
test on a part with no ``mc:AlternateContent``. Strict documents' namespace is not
matched, as CR-030's rules do not match it.

This module imports no model class (``tests/openpackaging/test_threads_and_import.py``).
"""

from __future__ import annotations

import copy
import re
from collections.abc import Iterable
from typing import Any

from docx4j_xsdata.formats.dataclass.parsers.skipped import SkippedNode
from lxml import etree

from docx4j_py.el_index import WML_ELEMENTS
from docx4j_py.openpackaging import breaks
from docx4j_py.openpackaging.mc_load_mode import branch_taken

__all__ = ["in_a_branch", "repair", "repair_bytes", "triggers"]

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_W_PREFIX = f"{{{_W}}}"
_ANY_W = f"{_W_PREFIX}*"


def _w(local: str) -> str:
    return f"{_W_PREFIX}{local}"


_P = _w("p")
_PPR = _w("pPr")
_PPR_CHANGE = _w("pPrChange")
_R = _w("r")
_SECT_PR = _w("sectPr")
_BRANCHES = frozenset((f"{{{_MC}}}Choice", f"{{{_MC}}}Fallback"))

#: The twenty hosts Word 365 reads through, each measured by a file of check 76.
_HOSTS = frozenset(
    _w(local)
    for local in (
        "p",
        "body",
        "tc",
        "r",
        "tbl",
        "tr",
        "sdtContent",
        "hyperlink",
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "hdr",
        "ftr",
        "comment",
        "fldSimple",
        "smartTag",
        "customXml",
        "footnote",
        "endnote",
    )
)
#: Six of the hosts' names stand for another element under these parents, which no file measures: a tracked change
#: marking a list number, a paragraph mark, a table row or a math control character, or holding a math run's content,
#: and a note naming a separator in the settings part. Every parent the schema gives those declarations, by way of its
#: groups and types, as ``test_the_other_elements_parents_are_the_schemas`` holds.
_ANOTHER_ELEMENT_UNDER = {
    _w("ins"): frozenset((_w("numPr"), _w("rPr"), _w("trPr"), f"{{{_MATH}}}ctrlPr", f"{{{_MATH}}}r")),
    _w("del"): frozenset((_w("rPr"), _w("trPr"), f"{{{_MATH}}}ctrlPr", f"{{{_MATH}}}r")),
    _w("moveFrom"): frozenset((_w("rPr"),)),
    _w("moveTo"): frozenset((_w("rPr"),)),
    _w("footnote"): frozenset((_w("footnotePr"),)),
    _w("endnote"): frozenset((_w("endnotePr"),)),
}
#: A declaration of the ``w:`` namespace, for the prefixes a part uses it under.
_W_PREFIX_DECLARED = re.compile(rb"""xmlns:([\w.-]+)\s*=\s*["']""" + re.escape(_W.encode()) + rb"""["']""")


def _unknown(tag: Any) -> bool:
    """Whether `tag` is a ``w:`` name the model doesn't hold."""
    return isinstance(tag, str) and tag.startswith(_W_PREFIX) and tag[len(_W_PREFIX) :] not in WML_ELEMENTS


def triggers(skipped: Iterable[Any]) -> bool:
    """Whether the fork's skipped report names a ``w:`` element the model has no name for.

    The fork reports the outermost element it skips, so this is true of a part
    any of whose unknown elements it met. It says nothing of the host: a part
    that holds one only where no rule takes it runs the pass and changes nothing.
    """
    return any(getattr(item, "kind", None) == "element" and _unknown(getattr(item, "qname", None)) for item in skipped)


def in_a_branch(data: bytes) -> bool:
    """Whether `data` may hold an unknown ``w:`` element in a kept ``mc`` branch, which the fork's report can't name.

    A wildcard holds a branch's unknown child, so no report lists it. This reads
    the bytes: a substring test where there's no ``mc:AlternateContent``, then the
    names used under each prefix of the ``w:`` namespace. It may answer yes for a
    part whose unknown name stands elsewhere, which costs a pass and changes nothing.
    """
    if b"AlternateContent" not in data:
        return False
    for prefix in set(_W_PREFIX_DECLARED.findall(data)):
        used = set(re.findall(rb"<" + re.escape(prefix) + rb":([\w.-]+)", data))
        if any(name.decode() not in WML_ELEMENTS for name in used):
            return True
    return False


def _holder(element: etree._Element) -> etree._Element | None:
    """The nearest ancestor that isn't itself an unknown ``w:`` element; None for a root."""
    for ancestor in element.iterancestors():
        if not _unknown(ancestor.tag):
            return ancestor
    return None


def _is_host(element: etree._Element) -> bool:
    if element.tag in _BRANCHES:
        return _taken_in_a_paragraph(element)
    if element.tag not in _HOSTS:
        return False
    parent = element.getparent()
    return parent is None or parent.tag not in _ANOTHER_ELEMENT_UNDER.get(element.tag, ())


def _taken_in_a_paragraph(branch: etree._Element) -> bool:
    """Whether `branch` is the one Word takes of a ``mc:AlternateContent`` standing in a ``w:p``, rule 2's measured place."""
    alternate = branch.getparent()
    holder = alternate.getparent() if alternate is not None else None
    return holder is not None and holder.tag == _P and branch_taken(alternate) is branch


def _paragraph_of(holder: etree._Element) -> etree._Element | None:
    """The paragraph rule 3 moves runs to from `holder`: None unless `holder` is one of its three places."""
    parent = holder.getparent()
    if parent is None:
        return None
    if holder.tag == _PPR and parent.tag == _PPR_CHANGE:
        parent = parent.getparent()  # the paragraph's own w:pPr
        return parent.getparent() if parent is not None and parent.tag == _PPR and _is_a_paragraph(parent.getparent()) else None
    if holder.tag == _SECT_PR and parent.tag == _PPR:
        return parent.getparent() if _is_a_paragraph(parent.getparent()) else None
    return parent if holder.tag == _PPR and parent.tag == _P else None


def _is_a_paragraph(element: etree._Element | None) -> bool:
    return element is not None and element.tag == _P


def _empties(element: etree._Element, runs: list[etree._Element]) -> bool:
    """Whether `element` is left holding no element once `runs` have gone. Text it holds is absent, as ever (REQ-314)."""
    return not any(isinstance(child.tag, str) and child not in runs for child in element)


def repair(root: etree._Element, dropped: list[SkippedNode] | None = None) -> int:
    """The rules applied to `root` in place; returns how many elements gave up what they held.

    `dropped`, given a list, receives a ``SkippedNode`` for each element gone, in
    document order, in the form the fork's report uses, its path as read: each
    unknown element, the ``w:p`` rule 4 ended its host with, and a ``w:sectPr`` rule 3
    took out. An element rule 3 moved a run out of and left holding something else
    stands, and the fork's next parse reports it, so it has no entry here.
    """
    through: list[etree._Element] = []
    ends: list[tuple[list[etree._Element], etree._Element]] = []
    to_head: list[tuple[etree._Element, list[etree._Element], etree._Element, bool]] = []
    sections: dict[etree._Element, None] = {}

    def report(element: etree._Element) -> None:
        if dropped is not None:
            dropped.append(breaks.skipped_node(element))

    for element in root.iter(_ANY_W):
        if not _unknown(element.tag):
            continue
        holder = _holder(element)
        if holder is None:
            continue
        if _is_host(holder):
            through.append(element)
            report(element)
            held = list(element.iterchildren(_P))
            if held and holder.tag == _P and _can_end(element, holder):
                ends.append((held, holder))
                report(held[0])
            continue
        paragraph = _paragraph_of(holder) if element.getparent() is holder else None
        runs = list(element.iterchildren(_R)) if paragraph is not None else []
        if runs:
            in_a_section = holder.tag == _SECT_PR
            goes = in_a_section or _empties(element, runs)
            to_head.append((element, runs, paragraph, goes))
            if goes:
                report(element)
            if in_a_section and holder not in sections:
                sections[holder] = None
                report(holder)

    for held, host in ends:
        _end_host(held, host)
    for element in through:
        _give_way(element)
    placed: dict[etree._Element, int] = {}
    for element, runs, paragraph, goes in to_head:
        _to_head(runs, paragraph, placed)
        if goes:
            _give_way(element)
    for section in sections:
        section.getparent().remove(section)
    return len(through) + len(to_head)


def _give_way(element: etree._Element) -> None:
    """`element` replaced in its parent by the elements it held, in order. Text it held directly goes (REQ-314)."""
    parent = element.getparent()
    if parent is None:  # pragma: no cover - _holder refuses a root
        return
    at = parent.index(element)
    tail = element.tail
    previous = element.getprevious()
    children = list(element)
    for offset, child in enumerate(children):
        child.tail = None
        parent.insert(at + offset, child)
    if tail:
        last = children[-1] if children else previous
        if last is not None:
            last.tail = (last.tail or "") + tail
        else:
            parent.text = (parent.text or "") + tail
    parent.remove(element)


def _can_end(element: etree._Element, host: etree._Element) -> bool:
    """Whether what `element` holds can end `host`: it comes last there, at any depth, and a run comes before it."""
    node = element
    while node is not host:
        if any(isinstance(sibling.tag, str) for sibling in node.itersiblings()):
            return False
        top, node = node, node.getparent()
    return any(isinstance(sibling.tag, str) and next(sibling.iter(_R), None) is not None for sibling in top.itersiblings(preceding=True))


def _end_host(held: list[etree._Element], host: etree._Element) -> None:
    """Rule 4: the first of `held` ends `host`, the rest stand after it, and what's left of `host` comes last.

    What's left is an empty paragraph with the host's properties: the host's own
    where a held ``w:pPr`` replaced them (``76t68``), and a copy where it kept them
    (``76t69``), less any ``w:sectPr``, which would be a second section break.
    """
    first, *further = held
    properties = first.find(_PPR)
    own = host.find(_PPR)
    left = etree.Element(_P)
    if own is not None and properties is not None:
        host.replace(own, properties)
        left.append(own)
    elif own is not None:
        kept = copy.deepcopy(own)
        for section in kept.findall(_SECT_PR):
            kept.remove(section)
        left.append(kept)
    elif properties is not None:
        host.insert(0, properties)
    _give_way(first)
    last = host
    for paragraph in further:
        last.addnext(paragraph)
        last = paragraph
    last.addnext(left)


def _to_head(runs: list[etree._Element], paragraph: etree._Element, placed: dict[etree._Element, int]) -> None:
    """`runs` moved to the head of `paragraph`, straight after its ``w:pPr`` and any already put there."""
    at = placed.get(paragraph)
    if at is None:
        at = paragraph.index(paragraph.find(_PPR)) + 1
    for run in runs:
        run.tail = None
        paragraph.insert(at, run)
        at += 1
    placed[paragraph] = at


def repair_bytes(data: bytes, dropped: list[SkippedNode] | None = None) -> tuple[bytes, int, int, int]:
    """A part's bytes as Word reads them: this pass and then CR-030's, on one tree.

    Args:
        data: the part's bytes.
        dropped: given a list, receives a ``SkippedNode`` for each element either
            pass took out, in the form the fork's report uses.

    Returns:
        ``(bytes, elements, breaks, text_elements)``: `elements` is
        :func:`repair`'s count and the last two are CR-030's. `data` itself where
        neither pass changed anything, else the rewritten part
        (:func:`~docx4j_py.openpackaging.breaks.write`).
    """
    root = breaks.parse(data)
    elements = repair(root, dropped)
    found, texts, changed = breaks.repair(root, dropped)
    if not elements and not changed:
        return data, 0, found, texts
    return breaks.write(root), elements, found, texts
