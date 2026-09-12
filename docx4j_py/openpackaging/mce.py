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
  *namespace* is one the generated model understands
  (:data:`docx4j_py.namespaces.UNDERSTOOD`), which is what ECMA-376 Part 3 says
  and what Word does.
* docx4j's ``mc:AlternateContent`` *inside a ``w:r``* is left alone, because its
  exporters resolve it later. This model parses it either way, so there is no
  reason to keep it: the preprocessor resolves every one.

The XSLT's other business --- the Strict-to-Transitional import and the
malformed-nesting repairs for Google Docs, pandoc and SSRS output --- is not
markup compatibility and is not here. The Strict import is
``RelationshipsPart.import_strict`` at the relationship level; the rest is a
later CR.
"""

from __future__ import annotations

from collections.abc import Iterable

from lxml import etree

__all__ = [
    "MCE_NS",
    "mce_preprocess",
    "mce_preprocess_bytes",
    "resolve_alternate_content",
]

MCE_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

_AC = f"{{{MCE_NS}}}AlternateContent"
_CHOICE = f"{{{MCE_NS}}}Choice"
_FALLBACK = f"{{{MCE_NS}}}Fallback"

_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)


def _understood() -> frozenset[str]:
    from docx4j_py.namespaces import UNDERSTOOD

    return frozenset(UNDERSTOOD.values())


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
        understood: the namespace URIs to count as understood; the generated
            model's namespaces by default.

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
