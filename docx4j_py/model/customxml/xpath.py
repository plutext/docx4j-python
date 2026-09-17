"""XPath over a custom XML part: lxml's, with Word's prefix-mapping strings.

CR-003 section 3.7. The one thing this port has that docx4j-core-ts had to
build is a complete XPath 1.0 engine: **lxml's is it**. There is no
``XPathEngine``, no optional dependency, nothing to warm and nothing
asynchronous --- ``select_nodes`` is a call.

What is here is the two things around it:

:func:`parse_prefix_mappings` / :func:`format_prefix_mappings`
    Word's ``xmlns:ns0='urn:invoice'`` string, the form ``w:dataBinding``
    carries in ``@w:prefixMappings`` and Office JS's ``namespaceMappings``
    argument takes, to and from a mapping.
:func:`canonical_xpath_of`
    the fully positional path of a node, ``/ns0:invoice[1]/ns0:total[1]``,
    with the prefixes numbered ``ns0``, ``ns1``, … in the order the namespaces
    are first met walking root to node --- which is what Word writes when it
    binds a control by picking a node in the XML Mapping pane.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, NamedTuple

if TYPE_CHECKING:  # pragma: no cover
    from lxml import etree

__all__ = [
    "CanonicalXPath",
    "canonical_xpath_of",
    "format_prefix_mappings",
    "parse_prefix_mappings",
    "select",
]

#: ``xmlns:ns0='urn:invoice'`` and ``xmlns="urn:invoice"``, either quote.
_MAPPING = re.compile(r"""xmlns(?::([\w.\-]+))?\s*=\s*(['"])(.*?)\2""")

#: lxml's node types, by the attribute a result carries.
_ELEMENT = 1
_ATTRIBUTE = 2


class CanonicalXPath(NamedTuple):
    """A node's canonical path and the prefix mappings it needs."""

    #: ``/ns0:invoice[1]/ns0:customer[1]/ns0:name[1]``.
    xpath: str
    #: ``xmlns:ns0='urn:invoice'``; ``""`` when no namespace is involved.
    prefix_mappings: str


def parse_prefix_mappings(mappings: str | None) -> dict[str, str]:
    """Word's ``xmlns:ns0='urn:invoice'`` string, as a prefix-to-URI mapping.

    Several declarations are separated by whitespace and either quote is
    accepted; a default declaration (``xmlns='…'``, no prefix) comes back under
    the key ``""``. ``None`` and ``""`` give an empty mapping, which is what
    ``samples/invoice2013.docx``'s twenty bindings carry.
    """
    if not mappings:
        return {}
    return {match.group(1) or "": match.group(3) for match in _MAPPING.finditer(mappings)}


def format_prefix_mappings(mappings: dict[str, str]) -> str:
    """The reverse: a mapping as the string Word writes, single-quoted.

    The order is the mapping's own, so a canonical path's ``ns0``, ``ns1``, …
    come out in the order they were allocated.
    """
    return " ".join(
        (f"xmlns='{uri}'" if not prefix else f"xmlns:{prefix}='{uri}'")
        for prefix, uri in mappings.items()
    )


def select(
    context: etree._Element,
    xpath: str,
    mappings: dict[str, str] | None = None,
) -> list:
    """Evaluate `xpath` with `context` as the context node.

    The empty prefix is dropped before the call: XPath 1.0 has no default
    namespace and lxml refuses a ``None`` key, so a mapping that carries one is
    used for its prefixed entries alone.

    Raises:
        BindingError: the expression is not valid XPath.
    """
    from docx4j_py.model.content.errors import BindingError

    namespaces = {prefix: uri for prefix, uri in (mappings or {}).items() if prefix and uri}
    try:
        found = context.xpath(xpath, namespaces=namespaces)
    except Exception as error:
        raise BindingError(
            f"{xpath!r} is not an XPath this part can evaluate: {error}",
            code="binding.bad_xpath",
            hint="XPath 1.0, with a prefix for every namespace you name",
        ) from error
    if isinstance(found, (str, float, bool)):
        return [found]
    return list(found)


# ---------------------------------------------------------------------------
# the canonical path of a node
# ---------------------------------------------------------------------------


def _split(tag: str) -> tuple[str, str]:
    """``{uri}local`` into ``(uri, local)``; ``("", local)`` for no namespace."""
    if tag.startswith("{"):
        uri, _, local = tag[1:].partition("}")
        return uri, local
    return "", tag


def _position_of(parent: etree._Element, node: etree._Element, tag: str) -> int:
    """The 1-based position of `node` among its siblings with the same tag."""
    position = 0
    for child in parent:
        if getattr(child, "tag", None) == tag:
            position += 1
            if child is node:
                return position
    return position or 1


def canonical_xpath_of(node: object) -> CanonicalXPath:
    """The fully positional path of an element or an attribute.

    ``/ns0:invoice[1]/ns0:customer[1]/ns0:name[1]``, with a predicate on every
    element step and none on an attribute step, and the prefixes numbered from
    ``ns0`` in the order the namespaces are first met walking root to node.
    Elements in no namespace keep their bare name, which is what
    ``samples/invoice2013.docx`` --- whose ``<invoice>`` has no namespace ---
    gets: ``/invoice[1]/customer[1]/contact[1]``, exactly the path Word wrote
    into its bindings.

    An attribute is passed as the ``(element, name)`` pair
    :class:`~docx4j_py.model.customxml.CustomXmlNode` holds it as.
    """
    attribute: str | None = None
    if isinstance(node, tuple):
        node, attribute = node  # type: ignore[assignment]

    prefixes: dict[str, str] = {}

    def qname(uri: str, local: str) -> str:
        if not uri:
            return local
        prefix = prefixes.get(uri)
        if prefix is None:
            prefix = f"ns{len(prefixes)}"
            prefixes[uri] = prefix
        return f"{prefix}:{local}"

    chain: list = []
    current = node
    while current is not None and getattr(current, "tag", None) is not None:
        chain.append(current)
        current = current.getparent()
    chain.reverse()

    steps: list[str] = []
    for element in chain:
        tag = element.tag
        if not isinstance(tag, str):  # a comment or a processing instruction
            continue
        uri, local = _split(tag)
        parent = element.getparent()
        index = 1 if parent is None else _position_of(parent, element, tag)
        steps.append(f"{qname(uri, local)}[{index}]")

    if attribute is not None:
        uri, local = _split(attribute)
        steps.append(f"@{qname(uri, local)}")

    mappings = {prefix: uri for uri, prefix in prefixes.items()}
    return CanonicalXPath("/" + "/".join(steps), format_prefix_mappings(mappings))
