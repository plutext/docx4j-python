"""Markup compatibility, the *preprocessing* half. CR-002 section 5.6.

Two halves exist and they are deliberately different:

**The default is lossless.** CR-002 decided question 1: the model round-trips
both branches of an ``mc:AlternateContent``, so the engine does not resolve
anything on read. What a consumer wants is the *view* Word takes, and that is
``mce="resolve"`` on the traversal functions (``docx4j_py.child.iter_children``
and CR-001's ``walk``/``find``/``text_of``), not a rewritten tree.

**``LoadOptions(mce_preprocess=True)`` gives docx4j's behaviour**: this module,
an lxml pass over a part's bytes before it is unmarshalled, equivalent to
docx4j's ``mc-preprocessor.xslt`` for the markup-compatibility rules. Each
``mc:AlternateContent`` is replaced, in place, by the children of the first
``mc:Choice`` whose ``Requires`` prefixes are all understood, else by the
children of its ``mc:Fallback``, else by nothing. Innermost first, because a
chosen branch can hold another ``mc:AlternateContent``.

Two differences from docx4j's XSLT, both deliberate and both recorded in CR-002
section 12:

* docx4j selects a choice by comparing the ``Requires`` **prefix tokens** against
  the ``docx4j.jaxb.mc.preferChoice`` property, which is empty by default --- so
  docx4j out of the box always takes the fallback. This resolves each prefix
  against the element's own namespace declarations and asks whether the
  *namespace* is in docx4j's published understood list
  (:func:`docx4j_py.child.understood_namespaces`, the list a reader of text and
  CR-012's ``word`` mode select by), which is what ECMA-376 Part 3 says and what
  Word does. Until CR-012 Phase A (Q8) it asked whether the generated model
  understood it (:data:`docx4j_py.namespaces.UNDERSTOOD`), which lacks
  ``w16sdtfl``, ``w16sdtdh``, ``w16du`` and ``w16cex`` and so took ``w16cid``
  where Word 365 takes ``w16sdtfl``; ``UNDERSTOOD`` stays the prefix table's own
  list.
* docx4j's ``mc:AlternateContent`` *inside a ``w:r``* is left alone, because its
  exporters resolve it later. This model parses it either way, so there is no
  reason to keep it: the preprocessor resolves every one.

The XSLT's other business --- the Strict-to-Transitional import and the
malformed-nesting repairs for Google Docs, pandoc and SSRS output --- is not
markup compatibility and is not here. The Strict import is
``RelationshipsPart.import_strict`` at the relationship level; the rest is a
later CR.

**A prefix ``mc:Ignorable`` names with no declaration** on the part's root (WD REQ-323, 0.103) names no namespace:
it is read as though the list didn't hold it, and :class:`UndeclaredPrefix` says so, once a part. A save that writes
the part anew drops such a prefix from the list (``XmlPart._marshal``); one that writes the source bytes keeps it.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Iterable
from typing import Any, ClassVar

from lxml import etree

__all__ = [
    "MCE_NS",
    "UndeclaredPrefix",
    "mce_preprocess",
    "mce_preprocess_bytes",
    "resolve_alternate_content",
]

MCE_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

_AC = f"{{{MCE_NS}}}AlternateContent"
_CHOICE = f"{{{MCE_NS}}}Choice"
_FALLBACK = f"{{{MCE_NS}}}Fallback"

_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)


@dataclasses.dataclass(frozen=True, slots=True)
class UndeclaredPrefix:
    """A part whose root's ``mc:Ignorable`` names a prefix with no declaration on it: ``mce.undeclared_prefix``.

    WD REQ-323 (0.103): such a prefix names no namespace, so it is read as though the list didn't hold it. Word repairs
    such a file and loses nothing (the ``undeclared-ignorable`` round, ``ui-01`` to ``ui-14``). A declaration on a
    descendant doesn't put it in scope. On ``part.warnings`` from the part's parse, once a part; the text model raises
    it for each part it reads. Measured on a part's root only, which is all this reads.

    Attributes:
        part_name: the part, the WD's selector.
        prefixes: each prefix, in the list's order, once.
    """

    code: ClassVar[str] = "mce.undeclared_prefix"

    part_name: str
    prefixes: tuple[str, ...]

    @property
    def detail(self) -> str:
        """The WD's ``detail``: the prefixes, separated by spaces."""
        return " ".join(self.prefixes)

    @property
    def message(self) -> str:
        """``/word/document.xml: mc:Ignorable names zz9, declared nowhere on its root; read as not listed (REQ-323)``."""
        return f"{self.part_name}: mc:Ignorable names {self.detail}, declared nowhere on its root; read as not listed (REQ-323)"

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: ``code``, the two fields and ``message``."""
        return {"code": self.code, "part_name": self.part_name, "prefixes": list(self.prefixes), "message": self.message}

    def __str__(self) -> str:
        """:attr:`message`."""
        return self.message


def undeclared_prefixes(ns_map: dict[str, str], ignorable: str | None) -> tuple[str, ...]:
    """The prefixes `ignorable`, an ``mc:Ignorable`` value, names that `ns_map` doesn't declare: each once, in order."""
    if not ignorable:
        return ()
    return tuple(dict.fromkeys(prefix for prefix in ignorable.split() if prefix not in ns_map))


def _understood() -> frozenset[str]:
    from docx4j_py.child import understood_namespaces

    return understood_namespaces()


def mce_preprocess_bytes(data: bytes, understood: Iterable[str] | None = None) -> bytes:
    """Resolve every ``mc:AlternateContent`` in a part and return the new bytes."""
    root = etree.fromstring(data, _PARSER)
    resolve_alternate_content(root, understood)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def mce_preprocess(root: etree._Element, understood: Iterable[str] | None = None) -> int:
    """Resolve every ``mc:AlternateContent`` under `root`, in place.

    Returns:
        How many were resolved.
    """
    return resolve_alternate_content(root, understood)


def resolve_alternate_content(
    root: etree._Element, understood: Iterable[str] | None = None
) -> int:
    """Replace every ``mc:AlternateContent`` below `root` with its chosen branch.

    Innermost first: a chosen branch may itself hold alternate content, and
    resolving the outer one first would leave the inner one behind.

    Args:
        root: the element to work under; it is not itself replaced.
        understood: the namespace URIs to count as understood; docx4j's
            published understood list by default (CR-012 Q8).

    Returns:
        The number of ``mc:AlternateContent`` elements resolved.
    """
    known = frozenset(understood) if understood is not None else _understood()
    count = 0
    # Depth first, deepest first, so that an inner AlternateContent is resolved
    # before the outer one that holds it moves its children.
    for node in reversed(list(root.iter(_AC))):
        _replace(node, known)
        count += 1
    return count


def _replace(node: etree._Element, known: frozenset[str]) -> None:
    parent = node.getparent()
    if parent is None:  # pragma: no cover - an AlternateContent root is not legal
        return
    branch = choose_branch(node, known)
    index = parent.index(node)
    tail = node.tail
    if branch is not None:
        children = list(branch)
        for offset, child in enumerate(children):
            parent.insert(index + offset, child)
        if children:
            # the removed element's tail belongs to whatever now stands last
            last = children[-1]
            last.tail = (last.tail or "") + (tail or "") if tail else last.tail
        elif branch.text and branch.text.strip():
            _append_text(parent, index, branch.text)
    parent.remove(node)


def _append_text(parent: etree._Element, index: int, text: str) -> None:
    if index == 0:
        parent.text = (parent.text or "") + text
    else:
        previous = parent[index - 1]
        previous.tail = (previous.tail or "") + text


def choose_branch(node: etree._Element, known: frozenset[str]) -> etree._Element | None:
    """The branch of an ``mc:AlternateContent`` element this consumer takes.

    ECMA-376 Part 3: the first ``mc:Choice`` all of whose ``Requires`` prefixes
    resolve to namespaces in `known`, else the ``mc:Fallback``, else None.
    """
    fallback = None
    for child in node:
        if not isinstance(child.tag, str):
            continue
        if child.tag == _CHOICE:
            if _requires_understood(child, known):
                return child
        elif child.tag == _FALLBACK and fallback is None:
            fallback = child
    return fallback


def _requires_understood(choice: etree._Element, known: frozenset[str]) -> bool:
    requires = (choice.get("Requires") or "").strip()
    if not requires:
        return False
    ns_map = choice.nsmap
    for prefix in requires.split():
        uri = ns_map.get(prefix)
        if uri is None or uri not in known:
            return False
    return True
