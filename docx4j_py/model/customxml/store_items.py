"""The custom XML part a binding is evaluated against, as Word repairs a stale ID. CR-014 phase 1c.

docx4j's ``StoreItemResolver`` (its CR-058 phase 2, ``35cdeaee6``; D4 and D5), held
to Word 365 on docx4j-portfolio's ``bind-plain`` probe and its two-parts and no-part
variants (22 and 23 October 2026). Where a ``w:storeItemID`` names no part, Word
evaluates the binding against the part whose root element matches the XPath's first
step, by namespace and local name, read with the binding's own prefix mappings. Of
several, it takes the first in part order whose XPath selects a node; its save writes
that part's ID into the control. Where none fits, the control keeps its surface and
its stale ID.

The binding passes and the text model go through :func:`resolve`:
:func:`~docx4j_py.model.customxml.bindings.apply_bindings`, which writes the part's
ID back and says so; ``update_from_content_controls``; and
:class:`~docx4j_py.model.text.bound_values.BoundValues`, which writes nothing.
:attr:`XmlMapping.custom_xml_part <docx4j_py.model.customxml.mapping.XmlMapping.custom_xml_part>`
stays strict: it is Office JS's read, and what Office JS gives for a stale ID is
unmeasured (Q11). :attr:`~docx4j_py.model.customxml.mapping.XmlMapping.resolved_part`
is this one's.

**Departures from docx4j's class** (CR-014 section 12): no ``strictStoreItemIds``,
since this package never failed on a stale ID (Q11); and no memo per document, since
which of two parts selects a node may change as the data does, and a stale ID is
rare. The three well-known store item IDs name no part here unless the package
carries one (CR-003 section 4) and are not stale: docx4j reaches its resolver past
them, and this answers None for them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx4j_py.model.customxml.parts import normalise_id
from docx4j_py.model.customxml.xpath import parse_prefix_mappings
from docx4j_py.namespaces import PREFIXES

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.customxml.parts import CustomXmlPart, CustomXmlPartCollection

__all__ = [
    "WELL_KNOWN_STORE_ITEM_IDS",
    "first_step",
    "resolve",
    "resolved_store_item_id",
]

#: The core, extended and cover-page properties' store item IDs, braces off and lower-cased (docx4j's
#: ``BindingHandler`` constants): the binding means Word's own property store, whether or not the package holds a part.
WELL_KNOWN_STORE_ITEM_IDS: frozenset[str] = frozenset(
    {
        "6c3c8bc8-f283-45ae-878a-bab7291924a1",
        "6668398d-a668-4e3e-a5eb-62b293d839f1",
        "55af091b-3c7a-41e3-b477-f2fdaa23cfda",
    }
)

#: What ends an XPath's first step.
_STEP_ENDS = "/[("


def first_step(xpath: str | None, prefix_mappings: str | None) -> tuple[str, str] | None:
    """The XPath's first step as ``(namespace URI, local name)``. docx4j ``firstStep``.

    ``/p:root[1]/p:b7[1]`` with ``xmlns:p='urn:x'`` is ``("urn:x", "root")``. A
    prefix the binding doesn't declare is looked up in docx4j's table, as the XPath
    would be read, and failing that is no namespace.

    Returns:
        The step, or None where the XPath doesn't start with a named one
        (``//x``, ``/*``, ``/p:*``, ``/.``, ``/@x``, an axis), so that every part fits.
    """
    if xpath is None:
        return None
    text = xpath.strip()
    if text.startswith("//"):
        return None
    text = text.removeprefix("/")
    end = next((i for i, c in enumerate(text) if c in _STEP_ENDS or c.isspace()), len(text))
    step = text[:end]
    if step in ("", "*", ".", "..") or "::" in step or step.startswith("@"):
        return None
    prefix, colon, local = step.partition(":")
    if not colon:
        prefix, local = "", step
    if local in ("", "*"):
        return None
    uri = (parse_prefix_mappings(prefix_mappings).get(prefix) or PREFIXES.get(prefix) or "") if prefix else ""
    return uri, local


def _root_step(part: CustomXmlPart) -> tuple[str, str] | None:
    """A part's root element as ``(namespace URI, local name)``, or None for a part that will not read."""
    try:
        root = part.tree
    except Exception:  # noqa: BLE001 - a part that will not parse fits nothing
        return None
    tag = getattr(root, "tag", None)
    if not isinstance(tag, str):
        return None
    uri, brace, local = tag[1:].partition("}") if tag.startswith("{") else ("", "", tag)
    return (uri, local) if brace or not tag.startswith("{") else None


def _selects_a_node(part: CustomXmlPart, xpath: str, prefix_mappings: str | None) -> bool:
    try:
        return part.select_single_node(xpath, prefix_mappings) is not None
    except Exception:  # noqa: BLE001 - an XPath that will not evaluate selects nothing
        return False


def _repair(collection: CustomXmlPartCollection, xpath: str, prefix_mappings: str | None) -> CustomXmlPart | None:
    """docx4j ``repair``: the part the XPath fits, or None. The first in part-name order among several, compared as text."""
    step = first_step(xpath, prefix_mappings)
    fitting = []
    for part in collection.items:
        root = _root_step(part)
        # a property store is no candidate: docx4j's are typed parts, not CustomXmlDataStorageParts
        if root is None or part.built_in or (step is not None and root != step):
            continue
        fitting.append(part)
    fitting.sort(key=lambda part: str(part.part.part_name))
    if len(fitting) < 2:
        return fitting[0] if fitting else None
    return next((part for part in fitting if _selects_a_node(part, xpath, prefix_mappings)), fitting[0])


def resolve(
    collection: CustomXmlPartCollection | None,
    store_item_id: str | None,
    xpath: str | None,
    prefix_mappings: str | None,
) -> CustomXmlPart | None:
    """The part a binding is evaluated against: the one its ID names, else the one its XPath fits. docx4j ``resolve``.

    Returns:
        The part, or None where the ID names none and none fits, or there is no
        ID at all, or it is a well-known store's.
    """
    if collection is None or not store_item_id:
        return None
    named = collection.get_item(store_item_id)
    if named is not None:
        return named
    if not xpath or normalise_id(store_item_id) in WELL_KNOWN_STORE_ITEM_IDS:
        return None
    return _repair(collection, xpath, prefix_mappings)


def resolved_store_item_id(
    collection: CustomXmlPartCollection | None,
    store_item_id: str | None,
    xpath: str | None,
    prefix_mappings: str | None,
) -> str | None:
    """The ID a binding should carry. docx4j ``resolvedStoreItemId``.

    The part's own, upper-cased, where the binding's names no part and one fits:
    what Word's save writes back. Else the binding's as it is.
    """
    if collection is None or not store_item_id or collection.get_item(store_item_id) is not None:
        return store_item_id
    part = resolve(collection, store_item_id, xpath, prefix_mappings)
    return part.id.upper() if part is not None else store_item_id
