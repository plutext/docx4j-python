"""An unknown ``w:`` element, read as Word 365 reads it. core-ts's CR-026 D7 and D8.

CR-037, on core-ts's Word check 76 (its CR-026 section 4.4), the WD's REQ-308 to
REQ-312. ``w:madeUp`` isn't a WordprocessingML element, and Word 365 doesn't treat
it as one thing. In a content host it takes the tags away and reads what they
held, at any depth, with no message. The model has no place for one, so the
fork's parse skips it and all it holds: every reader and every save went without
the words Word shows, with the loss in ``part.skipped``.

The rules, each decided on the tree as read and then applied:

1. an element in the ``w:`` namespace whose name the model doesn't hold
   (:data:`~docx4j_py.el_index.WML_ELEMENTS`), standing in one of the fifteen hosts
   measured, or in such an element standing in one, gives way to what it held, in
   order, with its text and tails (``76t`` to ``76t25``);
2. a ``w:r`` such an element holds directly, where the element stands directly in
   a paragraph's ``w:pPr``, goes to the head of the paragraph, straight after the
   ``w:pPr`` (``76t27``, core-ts's D8). What else the element holds stays in it, so
   the fork drops and reports that, as for a property (``76t5``).

Everything else stays as the parse leaves it: a property in ``w:pPr`` or ``w:rPr``
and anything in a ``w:t``, measured; a run in a ``w:rPr``, a file Word refuses
(``76t28``); and every host not measured, ``w:ftr``, ``w:endnote``, ``w:del``,
``w:sdtPr``, ``w:trPr``, ``w:tcPr``, ``w:tblPr``, ``w:sectPr`` and the rest. Adding a
host later is one line of :data:`_HOSTS`. An element of another namespace stops
the walk upward, since Word drops one with all it holds (``76k``), so a ``w:madeUp``
inside it isn't read through. A ``w:`` element with a name the model does hold is
never taken, though it's in the wrong place: the skipped report can't tell the
two apart, and that's what the table is for.

The pass doesn't look at what its host admits, which no Word run measures in a
content host: a ``w:p`` left in a ``w:p`` is the second parse's to skip and list,
as any misplaced element is.

Each element taken is one :class:`SkippedNode`, so ``part.skipped`` lists it and
``strict=True`` refuses, though no words go: the tags are lost whatever the words
do, and core-ts reports the same (CR-037 Q1).

:func:`repair_bytes` is the one repair of a part: this pass and then CR-030's, on
one tree, then one parse. The order is a design choice, not a measurement, as no
Word run holds the shape: a ``w:br`` such an element held directly is then rule 1's
to place. :func:`triggers` reads the fork's report, so a clean part pays the test
of an empty list. Strict documents' namespace is not matched, as CR-030's rules
do not match it.

This module imports no model class (``tests/openpackaging/test_threads_and_import.py``).
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from docx4j_xsdata.formats.dataclass.parsers.skipped import SkippedNode
from lxml import etree

from docx4j_py.el_index import WML_ELEMENTS
from docx4j_py.openpackaging import breaks

__all__ = ["repair", "repair_bytes", "triggers"]

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_W_PREFIX = f"{{{_W}}}"
_ANY_W = f"{_W_PREFIX}*"


def _w(local: str) -> str:
    return f"{_W_PREFIX}{local}"


_P = _w("p")
_PPR = _w("pPr")
_R = _w("r")

#: The fifteen hosts Word 365 reads through, each measured by a file of check 76.
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
        "hdr",
        "comment",
        "fldSimple",
        "smartTag",
        "customXml",
        "footnote",
    )
)
#: Two of the hosts' names also stand for a thing that holds no content: ``w:ins`` marking a list number, a paragraph
#: mark, a math control character, a table row or a math run, and ``w:footnote`` naming a separator in the settings
#: part. No file measures those, so each stays unmeasured. By parent, from the schema's declarations of the two names.
_NOT_CONTENT_UNDER = {
    _w("ins"): frozenset((_w("numPr"), _w("rPr"), _w("trPr"), f"{{{_MATH}}}r")),
    _w("footnote"): frozenset((_w("footnotePr"),)),
}
#: ``normalize-space``'s whitespace, as CR-030's.
_XML_WHITESPACE = " \t\r\n"


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


def _holder(element: etree._Element) -> etree._Element | None:
    """The nearest ancestor that isn't itself an unknown ``w:`` element; None for a root."""
    for ancestor in element.iterancestors():
        if not _unknown(ancestor.tag):
            return ancestor
    return None


def _is_host(element: etree._Element) -> bool:
    if element.tag not in _HOSTS:
        return False
    parent = element.getparent()
    return parent is None or parent.tag not in _NOT_CONTENT_UNDER.get(element.tag, ())


def _direct_runs(element: etree._Element, holder: etree._Element) -> list[etree._Element]:
    """The runs rule 2 moves out of `element`: none unless it stands directly in a paragraph's ``w:pPr``."""
    if holder.tag != _PPR or element.getparent() is not holder:
        return []
    paragraph = holder.getparent()
    if paragraph is None or paragraph.tag != _P:
        return []
    return list(element.iterchildren(_R))


def _empties(element: etree._Element, runs: list[etree._Element]) -> bool:
    """Whether `element` is left holding nothing once `runs` have gone."""
    if (element.text or "").strip(_XML_WHITESPACE):
        return False
    return not any(isinstance(child.tag, str) and child not in runs for child in element)


def repair(root: etree._Element, dropped: list[SkippedNode] | None = None) -> int:
    """The rules applied to `root` in place; returns how many elements gave up what they held.

    `dropped`, given a list, receives a ``SkippedNode`` for each element gone, in
    document order, in the form the fork's report uses, its path as read. An
    element rule 2 moved a run out of and left holding something else stands, and
    the fork's next parse reports it, so it has no entry here.
    """
    through: list[etree._Element] = []
    to_head: list[tuple[etree._Element, list[etree._Element], bool]] = []
    for element in root.iter(_ANY_W):
        if not _unknown(element.tag):
            continue
        holder = _holder(element)
        if holder is None:
            continue
        if _is_host(holder):
            through.append(element)
            if dropped is not None:
                dropped.append(breaks.skipped_node(element))
            continue
        runs = _direct_runs(element, holder)
        if runs:
            empties = _empties(element, runs)
            to_head.append((element, runs, empties))
            if empties and dropped is not None:
                dropped.append(breaks.skipped_node(element))

    for element in through:
        _give_way(element)
    placed: dict[etree._Element, int] = {}
    for element, runs, empties in to_head:
        _to_head(element, runs, placed)
        if empties:
            _give_way(element)
    return len(through) + len(to_head)


def _give_way(element: etree._Element) -> None:
    """`element` replaced in its parent by what it held, in order, with its text and tails."""
    parent = element.getparent()
    if parent is None:  # pragma: no cover - _holder refuses a root
        return
    at = parent.index(element)
    tail = element.tail
    previous = element.getprevious()
    if element.text:
        if previous is not None:
            previous.tail = (previous.tail or "") + element.text
        else:
            parent.text = (parent.text or "") + element.text
    children = list(element)
    for offset, child in enumerate(children):
        parent.insert(at + offset, child)
    if tail:
        last = children[-1] if children else previous
        if last is not None:
            last.tail = (last.tail or "") + tail
        else:
            parent.text = (parent.text or "") + tail
    parent.remove(element)


def _to_head(element: etree._Element, runs: list[etree._Element], placed: dict[etree._Element, int]) -> None:
    """`runs` moved to the head of `element`'s paragraph, after any already put there."""
    p_pr = element.getparent()
    paragraph = p_pr.getparent()
    at = placed.get(paragraph, paragraph.index(p_pr) + 1)
    for run in runs:
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
