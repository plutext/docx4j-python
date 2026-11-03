"""A ``w:br`` outside a ``w:r``, read as Word 365 reads it. docx4j CR-073's three rules.

CR-030, the port of docx4j CR-073 (``mc-preprocessor.xslt`` lines 830 to 951 at
``95c6b42d3``; core-ts CR-021 is the sibling). Word 365 opens a ``w:br`` found
directly in a ``w:p``, a ``w:hyperlink`` or another element that holds runs, one
inside a ``w:t``, ``w:instrText``, ``w:delText`` or ``w:delInstrText``, and one
directly in ``w:body``, with no repair, and reads each as a break at that point,
of its own type, all the text kept (docx4j's ``br-placement`` probe,
``tests/fixtures/br-placement/``). The model has no place for one there: the
fork's parse skips it (``part.skipped``), so every reader and every save went
without the break, and inside a ``w:t`` without the text after it. docx4j's load
survey found one in 206,351 of 3.7 million Common Crawl documents, 89% of them a
web-to-docx generator's blank line, ``<w:p><w:br/></w:p>``.

The rules, each decided on the tree as read and then applied, as the
stylesheet's templates match their input:

1. a ``w:br`` whose parent is a run container (``w:p``, ``w:hyperlink``,
   ``w:smartTag``, ``w:customXml``, ``w:fldSimple``, ``w:ins``, ``w:del``,
   ``w:moveFrom``, ``w:moveTo``, ``w:sdtContent``, ``w:dir``, ``w:bdo``) and
   whose nearest paragraph-or-block ancestor is a ``w:p`` goes in a run of its
   own, its attributes kept. The run takes the paragraph mark's run properties,
   without what only a mark has, where no run precedes the break in its
   paragraph, and none after one (docx4j's D1, Word's round 2);
2. a text element holding elements becomes its stretches of text, each under
   its own name, and its breaks, in order, in the same run; any other element in
   it gives its text, its own tag dropped. A stretch takes the ``xml:space`` of
   the nearest element at or above it that has one, and is kept where that is
   ``preserve`` or it holds more than XML whitespace;
3. a ``w:br`` directly in ``w:body`` goes into the next paragraph as its first
   run, after its ``w:pPr``, with that mark's properties; where no paragraph
   follows (or the next is one docx4j's nested-paragraph template splits), it is
   a paragraph of its own.

What triggers the pass is the fork's report (CR-030 Q1): :func:`triggers`
reads each skipped element's **path**, whose last step but one is its parent.
The report's ``parent_qname`` is a class's name for a choice class (``{w}RT``
for a ``w:t``), so it is not used (CR-030 section 2.3). A clean part pays a
scan of an empty list; an affected part is parsed, rewritten and parsed again
(:meth:`~docx4j_py.openpackaging.parts.xml_part.XmlPart._parse`), as docx4j's
retry after a failed unmarshal.

The counts are the stylesheet's (``count()`` at lines 834 to 837): every
``w:br`` whose parent is a run container or ``w:body``, and every text element
holding an element, at any depth. A ``w:br`` directly in a block-level
``w:sdtContent`` is counted and left where it is, so it stays in
``part.skipped``. Strict documents' namespace is not matched, as docx4j's rules
do not match it.

This module imports no model class (``tests/openpackaging/test_threads_and_import.py``).
"""

from __future__ import annotations

import copy
import dataclasses
import re
from collections.abc import Iterable
from typing import Any, ClassVar

from lxml import etree

__all__ = ["CODE", "BrOutsideRun", "repair", "repair_bytes", "triggers"]

#: The warning's code, agreed with core-ts through the portfolio (CR-030, settled).
CODE = "br.outside_run"

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(local: str) -> str:
    return f"{{{_W}}}{local}"


_BR = _w("br")
_P = _w("p")
_R = _w("r")
_BODY = _w("body")
_PPR = _w("pPr")
_RPR = _w("rPr")
_HYPERLINK = _w("hyperlink")
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

#: Rule 1's parents: the elements that hold runs. Named, since a math run (``m:r``) holds a ``w:br`` legally.
_CONTAINERS = frozenset(
    _w(local)
    for local in (
        "p",
        "hyperlink",
        "smartTag",
        "customXml",
        "fldSimple",
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "sdtContent",
        "dir",
        "bdo",
    )
)
#: The paragraph-or-block ancestors: rule 1 takes a ``w:br`` only where the nearest is a ``w:p``, so a block-level
#: ``w:sdtContent`` or ``w:customXml`` is not taken for an inline one.
_BLOCKS = frozenset(
    _w(local)
    for local in ("p", "body", "tc", "txbxContent", "hdr", "ftr", "footnote", "endnote", "comment", "docPartBody")
)
#: Rule 2's text elements.
_TEXTS = frozenset(_w(local) for local in ("t", "instrText", "delText", "delInstrText"))
_TEXT_TAGS = tuple(sorted(_TEXTS))
#: What only a paragraph mark's run properties hold, left out of a run's (``CT_ParaRPr``'s first and last).
_MARK_ONLY = frozenset(_w(local) for local in ("ins", "del", "moveFrom", "moveTo", "rPrChange"))
#: ``normalize-space``'s whitespace: XML's four characters, not ``str.strip``'s (U+00A0 is text).
_XML_WHITESPACE = " \t\r\n"

_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)

#: One step of a skipped item's path: ``{uri}local``, a bare local name, or an attribute's ``@`` form.
_STEP = re.compile(r"@?(?:\{[^}]*\})?[^/{}@]+")


# ---------------------------------------------------------------------------
# the record
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True, slots=True)
class BrOutsideRun:
    """A part read with a ``w:br`` outside a ``w:r``, as Word reads it: ``br.outside_run``, one per part.

    The code and the fields are agreed with docx4j-core-ts (CR-021's
    ``BrOutsideRunWarning``: ``partName``, ``breaks``, ``textElements``), and
    the message is docx4j's WARN after the part's name, less docx4j's own CR
    number (CR-030 section 3.2). The counts are the stylesheet's. Nothing is
    taken out, so there is no list of addresses: the source bytes stay in the
    store for any comparison.

    Attributes:
        part_name: the part, ``/word/document.xml``.
        breaks: the ``w:br`` found directly in a run container or in ``w:body``.
        text_elements: the ``w:t``, ``w:instrText``, ``w:delText`` and
            ``w:delInstrText`` found holding elements, at any depth.
    """

    code: ClassVar[str] = CODE

    part_name: str
    breaks: int
    text_elements: int

    @property
    def message(self) -> str:
        """``/word/document.xml: w:br outside a w:r: 8 outside a run, 1 text elements holding elements; read as ...``."""
        return (
            f"{self.part_name}: w:br outside a w:r: {self.breaks} outside a run, "
            f"{self.text_elements} text elements holding elements; read as Word reads them"
        )

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view, for the MCP server: ``code``, the three fields and ``message``."""
        return {
            "code": self.code,
            "part_name": self.part_name,
            "breaks": self.breaks,
            "text_elements": self.text_elements,
            "message": self.message,
        }

    def __str__(self) -> str:
        """:attr:`message`."""
        return self.message


# ---------------------------------------------------------------------------
# the trigger
# ---------------------------------------------------------------------------


def _parent_step(path: str) -> str | None:
    """The last step but one of a skipped item's path: its parent element's ``{uri}local``."""
    steps = _STEP.findall(path)
    return steps[-2] if len(steps) >= 2 else None


def triggers(skipped: Iterable[Any]) -> bool:
    """Whether the fork's skipped report names a shape the rules repair (CR-030 section 4.1).

    True where a skipped element is a ``w:br`` whose parent is a run container
    or ``w:body``, or any element whose parent is a text element. The parent is
    read from the item's ``path``, never its ``parent_qname`` (section 2.3). An
    attribute never triggers: rule 2 matches a text element's element children.
    """
    for item in skipped:
        if getattr(item, "kind", None) != "element":
            continue
        parent = _parent_step(getattr(item, "path", "") or "")
        if parent in _TEXTS:
            return True
        if getattr(item, "qname", None) == _BR and (parent in _CONTAINERS or parent == _BODY):
            return True
    return False


# ---------------------------------------------------------------------------
# the pass
# ---------------------------------------------------------------------------


def repair_bytes(data: bytes) -> tuple[bytes, int, int]:
    """A part's bytes with the rules applied, and the stylesheet's two counts.

    Returns:
        ``(bytes, breaks, text_elements)``: `data` itself where the rules
        changed nothing (no shape, or only a ``w:br`` no rule takes, such as
        one in a block-level ``w:sdtContent``), else the rewritten part, written
        as :func:`~docx4j_py.openpackaging.mc_load_mode.splice_word` writes it.
    """
    root = etree.fromstring(data, _PARSER)
    breaks, texts, changed = _repair(root)
    if not changed:
        return data, breaks, texts
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), breaks, texts


def repair(root: etree._Element) -> tuple[int, int]:
    """The rules applied to `root` in place; returns the stylesheet's two counts, ``(breaks, text_elements)``."""
    breaks, texts, _changed = _repair(root)
    return breaks, texts


def _repair(root: etree._Element) -> tuple[int, int, bool]:
    breaks, texts = _counts(root)
    if breaks + texts == 0:
        return 0, 0, False

    # decided on the tree as read; a w:br inside a text element is rule 2's
    in_runs: list[tuple[etree._Element, etree._Element | None]] = []
    in_body: list[tuple[etree._Element, etree._Element | None]] = []
    for br in root.iter(_BR):
        parent = br.getparent()
        if parent is None or (parent.tag not in _CONTAINERS and parent.tag != _BODY):
            continue
        if _inside_text(br):
            continue
        if parent.tag == _BODY:
            following = _next_block(br)
            in_body.append((br, following if following is not None and _takes_body_breaks(following) else None))
            continue
        paragraph = _paragraph_of(br)
        if paragraph is None:
            continue  # a block-level w:sdtContent or w:customXml: no rule takes it
        in_runs.append((br, None if _run_precedes(br, paragraph) else _mark_run_properties(paragraph)))
    split = [t for t in root.iter(*_TEXT_TAGS) if _holds_element(t) and not _inside_text(t)]
    marks = {p: _mark_run_properties(p) for _br, p in in_body if p is not None}

    # (1)
    for br, run_properties in in_runs:
        _replace(br, [_run(run_properties, br)])
    # (2)
    for text in split:
        pieces: list[etree._Element] = []
        _split(text, etree.QName(text).localname, pieces)
        _replace(text, pieces)
    # (3): each paragraph's breaks after its w:pPr, in order
    placed: dict[etree._Element, int] = {}
    for br, paragraph in in_body:
        if paragraph is None:
            own = etree.Element(_P)
            own.append(_run(None, br))
            _replace(br, [own])
            continue
        properties = marks[paragraph]
        run = _run(copy.deepcopy(properties) if properties is not None else None, br)
        if paragraph not in placed:
            p_pr = paragraph.find(_PPR)
            placed[paragraph] = paragraph.index(p_pr) + 1 if p_pr is not None else 0
        paragraph.insert(placed[paragraph], run)
        placed[paragraph] += 1
        _replace(br, [])
    changed = bool(in_runs or split or in_body)
    return breaks, texts, changed


def _counts(root: etree._Element) -> tuple[int, int]:
    """The stylesheet's ``$breaks`` and ``$texts``."""
    breaks = 0
    for br in root.iter(_BR):
        parent = br.getparent()
        if parent is not None and (parent.tag in _CONTAINERS or parent.tag == _BODY):
            breaks += 1
    texts = sum(1 for t in root.iter(*_TEXT_TAGS) if _holds_element(t))
    return breaks, texts


def _holds_element(node: etree._Element) -> bool:
    return any(isinstance(child.tag, str) for child in node)


def _inside_text(node: etree._Element) -> bool:
    return any(ancestor.tag in _TEXTS for ancestor in node.iterancestors())


def _paragraph_of(br: etree._Element) -> etree._Element | None:
    """The break's paragraph, where its nearest paragraph-or-block ancestor is one."""
    for ancestor in br.iterancestors():
        if ancestor.tag in _BLOCKS:
            return ancestor if ancestor.tag == _P else None
    return None


def _run_precedes(br: etree._Element, paragraph: etree._Element) -> bool:
    """Whether a run precedes the break in its paragraph: on its path from the paragraph, an earlier sibling is or holds one."""
    node: etree._Element | None = br
    while node is not None and node is not paragraph:
        for sibling in node.itersiblings(preceding=True):
            if isinstance(sibling.tag, str) and next(sibling.iter(_R), None) is not None:
                return True
        node = node.getparent()
    return False


def _mark_run_properties(paragraph: etree._Element) -> etree._Element | None:
    """The paragraph mark's run properties as a run's, a copy: without what only a mark has; None where none is left."""
    kept = [
        child
        for p_pr in paragraph.iterchildren(_PPR)
        for r_pr in p_pr.iterchildren(_RPR)
        for child in r_pr
        if isinstance(child.tag, str) and child.tag not in _MARK_ONLY
    ]
    if not kept:
        return None
    out = etree.Element(_RPR)
    for child in kept:
        property_ = copy.deepcopy(child)
        property_.tail = None
        out.append(property_)
    return out


def _next_block(br: etree._Element) -> etree._Element | None:
    """The first element after a body-level break that is not one."""
    for sibling in br.itersiblings():
        if isinstance(sibling.tag, str) and sibling.tag != _BR:
            return sibling
    return None


def _takes_body_breaks(element: etree._Element) -> bool:
    """A ``w:p`` the body's breaks before it go into: not one docx4j's nested-paragraph template splits.

    That template (``w:hyperlink/w:p``, ``w:r/w:p``, ``w:hyperlink/w:r/w:p``,
    ``w:r/w:r/w:p``) is not ported here (CR-002 section 12.3 item 7); the test is
    copied so that a later port of it gives docx4j's result.
    """
    if element.tag != _P:
        return False
    for child in element:
        if child.tag not in (_HYPERLINK, _R):
            continue
        for grandchild in child:
            if grandchild.tag == _P:
                return False
            if grandchild.tag == _R and grandchild.find(_P) is not None:
                return False
    return True


def _bare_break(br: etree._Element) -> etree._Element:
    """The break as the stylesheet writes it: a new ``w:br`` with its attributes, no content."""
    out = etree.Element(_BR)
    for key, value in br.attrib.items():
        out.set(key, value)
    return out


def _run(run_properties: etree._Element | None, br: etree._Element) -> etree._Element:
    run = etree.Element(_R)
    if run_properties is not None:
        run.append(run_properties)
    run.append(_bare_break(br))
    return run


def _replace(node: etree._Element, replacements: list[etree._Element]) -> None:
    """`node` replaced by `replacements` in its parent, its tail kept after them."""
    parent = node.getparent()
    if parent is None:  # pragma: no cover - a root is never a break or a text element in a part
        return
    index = parent.index(node)
    tail = node.tail
    for offset, replacement in enumerate(replacements):
        parent.insert(index + offset, replacement)
    if replacements:
        replacements[-1].tail = tail
    elif tail:
        previous = node.getprevious()
        if previous is not None:
            previous.tail = (previous.tail or "") + tail
        else:
            parent.text = (parent.text or "") + tail
    parent.remove(node)


def _space_of(holder: etree._Element) -> str:
    """``string(ancestor-or-self::*[@xml:space][1]/@xml:space)`` from a text node's parent: empty where none has one."""
    node: etree._Element | None = holder
    while node is not None:
        value = node.get(_XML_SPACE)
        if value is not None:
            return value
        node = node.getparent()
    return ""


def _split(holder: etree._Element, name: str, out: list[etree._Element]) -> None:
    """Rule 2's ``split-text`` mode: `holder`'s text under `name`, its ``w:br`` as breaks, any other element's text."""
    _piece(holder.text, holder, name, out)
    for child in holder:
        if isinstance(child.tag, str):
            if child.tag == _BR:
                out.append(_bare_break(child))
            else:
                _split(child, name, out)
        # a comment or processing instruction gives nothing, the text after it its own stretch
        _piece(child.tail, holder, name, out)


def _piece(text: str | None, holder: etree._Element, name: str, out: list[etree._Element]) -> None:
    if not text:
        return
    space = _space_of(holder)
    if space != "preserve" and not text.strip(_XML_WHITESPACE):
        return
    piece = etree.Element(_w(name))
    if space:
        piece.set(_XML_SPACE, space)
    piece.text = text
    out.append(piece)
