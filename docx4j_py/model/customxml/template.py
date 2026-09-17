"""``describe()`` and ``fill()``: the template story, as a library call.

CR-003 section 3.7. docx4j-mcp's two template tools have a library half each,
and this is it:

``pkg.custom_xml_parts.describe()``
    what the document **wants**: every XPath it binds, the kind of control that
    shows it, the repeat it is inside if it is inside one, the part it lives in
    and the value the data holds now. A frozen :class:`Skeleton` with
    ``to_dict()`` and ``to_json()``, which is what a tool result is.
``pkg.custom_xml_parts.fill(data)``
    a ``dict`` keyed by XPath, by ``w:tag`` or by ``w:alias``, or the XML of a
    whole part as a string; the nodes are set and
    :func:`~docx4j_py.model.customxml.bindings.apply_bindings` is run, so the
    document shows the new values without waiting for Word to open it. A key
    that names a **repeat** takes a *list*: one data node per entry, cloned from
    the template node the document already has (CR-003 section 17.10).

**What Word does with a repeat, and what this does.** Word expands a bound
``w15:repeatingSection`` when it *opens* the document: it clones the section's
one ``w15:repeatingSectionItem`` once per node the binding's XPath selects, and
reconciles the items to the nodes. So the number of rows a reader sees is the
number of nodes in the **data**, and a filled document carries one template item
however many items it will show. ``fill()`` therefore writes the *nodes* --- and
leaves the expansion, which is Word's, to Word.

**Reading costs nothing.** ``describe()`` reads the main document part through
lxml, from the bytes the part would be saved as, unless the part is already
unmarshalled --- in which case the views are there and are used, and report the
addresses as well.

**Out of scope, deliberately** (CR-003 section 17): OpenDoPE's conditions
(``od:condition``), its repeats-as-a-section-per-node (``od:repeat``) and its
``od:Handler`` extensions. An OpenDoPE **XPaths part** is *read* when the
document has one, because that is where docx4j-mcp's ``describe_template``
looks for the question names; ``samples/invoice2013.docx`` has none, so its
skeleton comes from ``w:dataBinding`` alone.
"""

from __future__ import annotations

import copy
import json
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from lxml import etree

from docx4j_py.model.content.errors import BindingError
from docx4j_py.model.customxml.bindings import (
    BindingResult,
    apply_bindings,
    bound_bodies,
    controls_of,
)
from docx4j_py.model.customxml.nodes import CustomXmlNode
from docx4j_py.model.customxml.parts import custom_xml_parts_of
from docx4j_py.model.customxml.xpath import canonical_xpath_of
from docx4j_py.wml.sdt import KIND_BY_QNAME, W15_NS, W_NS

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.customxml.parts import CustomXmlPart

__all__ = [
    "MAX_VALUE_CHARS",
    "OPENDOPE_XPATHS_NS",
    "BindingInfo",
    "FillEntry",
    "FillResult",
    "RepeatInfo",
    "Skeleton",
    "TemplatePart",
    "describe_template",
    "fill_template",
]

#: The OpenDoPE XPaths part's namespace. A document that has one names its
#: bindings there as well; docx4j's ``XPathsPart`` is the Java reader.
OPENDOPE_XPATHS_NS = "http://opendope.org/xpaths"

#: How much of a value ``describe()`` reports. A picture binding's node holds
#: the whole image as base64 --- 28 KB in ``samples/invoice2013.docx`` --- and a
#: skeleton is a tool result, so a longer value is cut and marked with an
#: ellipsis (CR-003 section 3.4's budgets, section 17).
MAX_VALUE_CHARS = 200

_W = f"{{{W_NS}}}"
_W15 = f"{{{W15_NS}}}"

#: A ``fill()`` entry key that is a plain child element name rather than a
#: relative XPath: no step separator, no predicate, no axis, no attribute.
_PLAIN_NAME = re.compile(r"^[^/\[\]@():*\s]+$")


# ---------------------------------------------------------------------------
# the skeleton
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class BindingInfo:
    """One XPath the document binds, and what it is for. CR-003 section 3.1."""

    #: The bound XPath, as the control stores it.
    xpath: str
    #: The ``xmlns:ns0='…'`` string it needs, or ``""``.
    prefix_mappings: str
    #: The ``w:storeItemID`` of the part it reads.
    store_item_id: str
    #: The Office JS ``Word.ContentControlType`` of the control that shows it.
    kind: str
    #: ``w:tag``, the machine-readable name.
    tag: str
    #: ``w:alias``, which Word's dialog calls the title.
    title: str
    #: The control's ordinal address, or ``""`` when the part was read as bytes.
    address: str = ""
    #: The value the data holds now, or None when the XPath does not resolve.
    value: str | None = None
    #: The XPath of the repeating section this binding is inside, or None.
    repeat: str | None = None
    #: The part name the value lives in, or ``""``.
    part_name: str = ""
    #: The OpenDoPE ``xpath/@id`` this binding came from, when it did.
    open_dope_id: str = ""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        out: dict[str, Any] = {
            "xpath": self.xpath,
            "kind": self.kind,
            "tag": self.tag,
            "title": self.title,
            "value": self.value,
        }
        if self.prefix_mappings:
            out["prefix_mappings"] = self.prefix_mappings
        if self.store_item_id:
            out["store_item_id"] = self.store_item_id
        if self.address:
            out["address"] = self.address
        if self.repeat:
            out["repeat"] = self.repeat
        if self.part_name:
            out["part_name"] = self.part_name
        if self.open_dope_id:
            out["open_dope_id"] = self.open_dope_id
        return out


@dataclass(frozen=True, slots=True)
class RepeatInfo:
    """One repeating section, and what an entry of it takes. CR-003 section 17.10.

    Word expands a bound ``w15:repeatingSection`` to one section item per node
    the binding selects when it opens the document, so :attr:`count` is how many
    rows a reader will see and :attr:`fields` is what a ``fill()`` list entry may
    name --- the template node's own child elements. A repeat whose node carries
    **no** child elements (``samples/invoice2013.docx``'s ``note``) has no fields
    and takes a list of plain values.
    """

    #: The XPath the repeating section binds, as the control stores it.
    xpath: str
    #: The child element names one entry may set, in the template node's order.
    fields: tuple[str, ...] = ()
    #: How many nodes the data holds now: how many items Word will show.
    count: int = 0
    #: The ``xmlns:ns0='…'`` string :attr:`xpath` needs, or ``""``.
    prefix_mappings: str = ""
    #: The ``w:storeItemID`` of the part the nodes live in.
    store_item_id: str = ""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        out: dict[str, Any] = {
            "xpath": self.xpath,
            "fields": list(self.fields),
            "count": self.count,
        }
        if self.prefix_mappings:
            out["prefix_mappings"] = self.prefix_mappings
        if self.store_item_id:
            out["store_item_id"] = self.store_item_id
        return out

    def __repr__(self) -> str:
        """``<RepeatInfo /invoice[1]/lines[1]/lineitem[1] 2 items (productcode, …)>``."""
        fields = ", ".join(self.fields)
        return f"<RepeatInfo {self.xpath} {self.count} item(s) ({fields})>"


@dataclass(frozen=True, slots=True)
class TemplatePart:
    """One custom XML part the template reads from."""

    #: The ``ds:itemID``.
    id: str
    #: The document element's namespace, or ``""``.
    namespace_uri: str
    #: ``/customXml/item1.xml``.
    part_name: str
    #: Whether this is one of Word's own property stores.
    built_in: bool = False

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        return {
            "id": self.id,
            "namespace_uri": self.namespace_uri,
            "part_name": self.part_name,
            "built_in": self.built_in,
        }


@dataclass(frozen=True, slots=True)
class Skeleton:
    """What a template wants filling in. docx4j-mcp's ``describe_template``."""

    #: Every binding the document carries, in document order.
    bindings: tuple[BindingInfo, ...] = ()
    #: The repeating sections, in document order, with the fields an entry takes
    #: and how many nodes the data holds now (CR-003 section 17.10).
    repeats: tuple[RepeatInfo, ...] = ()
    #: The custom XML parts the bindings read.
    parts: tuple[TemplatePart, ...] = ()
    #: Whether an OpenDoPE XPaths part was found and read.
    open_dope: bool = False
    #: What could not be reported, and why.
    warnings: tuple[str, ...] = field(default=())

    @property
    def resolved(self) -> int:
        """How many bindings find a node in the data as it stands."""
        return sum(1 for binding in self.bindings if binding.value is not None)

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: what a tool result returns."""
        return {
            "bindings": [binding.to_dict() for binding in self.bindings],
            "repeats": [repeat.to_dict() for repeat in self.repeats],
            "parts": [part.to_dict() for part in self.parts],
            "open_dope": self.open_dope,
            "resolved": self.resolved,
            "warnings": list(self.warnings),
        }

    def to_json(self, **options: Any) -> str:
        """:meth:`to_dict` as JSON."""
        return json.dumps(self.to_dict(), **options)

    def __len__(self) -> int:
        """How many bindings there are."""
        return len(self.bindings)

    def __repr__(self) -> str:
        """``<Skeleton 20 bindings, 2 repeats, 1 part>``."""
        return (
            f"<Skeleton {len(self.bindings)} bindings, "
            f"{len(self.repeats)} repeats, {len(self.parts)} part"
            f"{'' if len(self.parts) == 1 else 's'}>"
        )


# ---------------------------------------------------------------------------
# describe()
# ---------------------------------------------------------------------------


def describe_template(package: Any, *, max_value_chars: int | None = MAX_VALUE_CHARS) -> Skeleton:
    """The skeleton of a bound document. See the module docstring.

    Args:
        max_value_chars: how much of each current value to report; None for all
            of it. The default keeps a picture binding's base64 out of a tool
            result.
    """
    collection = custom_xml_parts_of(package)
    parts = tuple(
        TemplatePart(
            id=part.id,
            namespace_uri=part.namespace_uri,
            part_name=str(part.part.part_name),
            built_in=part.built_in,
        )
        for part in collection.items
    )
    main = getattr(package, "main_document_part", None)
    warnings: list[str] = []
    if main is not None and getattr(main, "is_unmarshalled", False):
        bindings, repeats = _from_views(package, collection, max_value_chars)
    else:
        bindings, repeats = _from_bytes(main, collection, max_value_chars)
        if main is not None:
            warnings.append(
                "the main document part was read as bytes, so the bindings carry no address; "
                "read pkg.body first to get them"
            )
    open_dope_bindings = _open_dope(collection, max_value_chars)
    if open_dope_bindings:
        known = {binding.xpath for binding in bindings}
        bindings = bindings + tuple(
            binding for binding in open_dope_bindings if binding.xpath not in known
        )
    return Skeleton(
        bindings=bindings,
        repeats=_repeat_infos(collection, repeats),
        parts=parts,
        open_dope=bool(open_dope_bindings),
        warnings=tuple(warnings),
    )


def _local(tag: Any) -> str:
    """``{uri}local`` into ``local``; ``""`` for a comment or a PI."""
    if not isinstance(tag, str):
        return ""
    return tag.rpartition("}")[2] if tag.startswith("{") else tag


def _repeat_infos(
    collection: Any, raw: tuple[tuple[str, str, str], ...]
) -> tuple[RepeatInfo, ...]:
    """What a repeat's list entry takes, read from the template node itself.

    The fields are the child element names of the node the repeat's XPath
    selects --- which is the node Word clones --- and the count is how many
    siblings of that node the data holds, which is how many items Word will show
    (CR-003 section 17.10).
    """
    out: list[RepeatInfo] = []
    for xpath, mappings, store_item_id in raw:
        fields: tuple[str, ...] = ()
        count = 0
        element = _repeat_node(collection, xpath, mappings, store_item_id)
        if element is not None:
            fields = tuple(
                dict.fromkeys(_local(child.tag) for child in element if isinstance(child.tag, str))
            )
            parent = element.getparent()
            count = (
                sum(1 for sibling in parent if sibling.tag == element.tag)
                if parent is not None
                else 1
            )
        out.append(
            RepeatInfo(
                xpath=xpath,
                fields=fields,
                count=count,
                prefix_mappings=mappings,
                store_item_id=store_item_id,
            )
        )
    return tuple(out)


def _repeat_node(collection: Any, xpath: str, mappings: str, store_item_id: str) -> Any:
    """The lxml element a repeat's XPath selects, or None."""
    part = collection.get_item(store_item_id) if store_item_id else None
    if part is None:
        return None
    try:
        node = part.select_single_node(xpath, mappings)
    except BindingError:
        return None
    if node is None or node.node_type != "Element":
        return None
    return node.element


def _value_for(
    collection: Any, store_item_id: str, xpath: str, mappings: str, limit: int | None
) -> tuple[Any, str]:
    """The value an XPath reads now, cut to `limit`, and the part it read it from."""
    part = collection.get_item(store_item_id) if store_item_id else None
    if part is None:
        return None, ""
    try:
        node = part.select_single_node(xpath, mappings)
    except BindingError:
        return None, str(part.part.part_name)
    if node is None:
        return None, str(part.part.part_name)
    text = node.text
    if limit is not None and len(text) > limit:
        text = text[:limit] + "\u2026"
    return text, str(part.part.part_name)


def _from_views(
    package: Any, collection: Any, limit: int | None = MAX_VALUE_CHARS
) -> tuple[tuple[BindingInfo, ...], tuple[tuple[str, str, str], ...]]:
    """The bindings read through the content-API views, addresses and all."""
    bindings: list[BindingInfo] = []
    repeats: list[tuple[str, str, str]] = []
    for control in controls_of(package):
        mapping = control.xml_mapping
        if control.type == "RepeatingSection" and mapping.is_mapped:
            repeats.append((mapping.xpath, mapping.prefix_mappings, mapping.store_item_id))
        if not mapping.is_mapped:
            continue
        value, part_name = _value_for(
            collection, mapping.store_item_id, mapping.xpath, mapping.prefix_mappings, limit
        )
        bindings.append(
            BindingInfo(
                xpath=mapping.xpath,
                prefix_mappings=mapping.prefix_mappings,
                store_item_id=mapping.store_item_id,
                kind=control.type,
                tag=control.tag,
                title=control.title,
                address=control.address,
                value=value,
                repeat=_enclosing_repeat(control),
                part_name=part_name,
            )
        )
    return tuple(bindings), tuple(repeats)


def _enclosing_repeat(control: Any) -> str | None:
    """The XPath of the repeating section a control is inside, or None."""
    from docx4j_py.model.content.controls import parent_control_of

    current = parent_control_of(control.element, control.parent_body)
    for _ in range(32):
        if current is None:
            return None
        if current.type == "RepeatingSection":
            return current.xml_mapping.xpath or None
        current = parent_control_of(current.element, current.parent_body)
    return None


def _from_bytes(
    main: Any, collection: Any, limit: int | None = MAX_VALUE_CHARS
) -> tuple[tuple[BindingInfo, ...], tuple[tuple[str, str, str], ...]]:
    """The bindings read with lxml, from the bytes the part would be saved as.

    Nothing is unmarshalled, so a document only ever *described* goes out byte
    for byte (CR-003 section 3.7).
    """
    if main is None:
        return (), ()
    try:
        root = etree.fromstring(main.xml)
    except Exception:  # noqa: BLE001 - a part that will not parse describes nothing
        return (), ()

    bindings: list[BindingInfo] = []
    repeats: list[tuple[str, str, str]] = []
    for pr in root.iter(f"{_W}sdtPr"):
        binding = _binding_element(pr)
        kind = _kind_of(pr)
        xpath = binding.get(f"{_W}xpath") if binding is not None else None
        mappings = (binding.get(f"{_W}prefixMappings") or "") if binding is not None else ""
        store_item_id = (binding.get(f"{_W}storeItemID") or "") if binding is not None else ""
        if kind == "RepeatingSection" and xpath:
            repeats.append((xpath, mappings, store_item_id))
        if binding is None or not xpath:
            continue
        value, part_name = _value_for(collection, store_item_id, xpath, mappings, limit)
        bindings.append(
            BindingInfo(
                xpath=xpath,
                prefix_mappings=mappings,
                store_item_id=store_item_id,
                kind=kind,
                tag=_val(pr, "tag"),
                title=_val(pr, "alias"),
                value=value,
                repeat=_enclosing_repeat_element(pr),
                part_name=part_name,
            )
        )
    return tuple(bindings), tuple(repeats)


def _binding_element(pr: Any) -> Any:
    """``w:dataBinding`` or ``w15:dataBinding``, whichever the control carries."""
    for name in (f"{_W}dataBinding", f"{_W15}dataBinding"):
        found = pr.find(name)
        if found is not None:
            return found
    return None


def _kind_of(pr: Any) -> str:
    """The Office JS kind a ``w:sdtPr`` element describes, read with lxml."""
    for child in pr:
        kind = KIND_BY_QNAME.get(child.tag) if isinstance(child.tag, str) else None
        if kind is not None:
            return kind
    return "RichText"


def _val(pr: Any, local_name: str) -> str:
    """``w:sdtPr/w:<local_name>/@w:val``, or ``""``."""
    found = pr.find(f"{_W}{local_name}")
    return (found.get(f"{_W}val") or "") if found is not None else ""


def _enclosing_repeat_element(pr: Any) -> str | None:
    """The repeating section above a ``w:sdtPr``, read with lxml."""
    current = pr.getparent()
    while current is not None:
        if isinstance(current.tag, str) and current.tag == f"{_W}sdt":
            other = current.find(f"{_W}sdtPr")
            if other is not None and other is not pr and _kind_of(other) == "RepeatingSection":
                binding = _binding_element(other)
                if binding is not None:
                    return binding.get(f"{_W}xpath")
                return None
        current = current.getparent()
    return None


def _open_dope(collection: Any, limit: int | None = MAX_VALUE_CHARS) -> tuple[BindingInfo, ...]:
    """The bindings an OpenDoPE XPaths part declares, when the document has one."""
    out: list[BindingInfo] = []
    for part in collection.get_by_namespace(OPENDOPE_XPATHS_NS):
        root = part.document_element
        if root is None:
            continue
        for entry in root.element.iter(f"{{{OPENDOPE_XPATHS_NS}}}xpath"):
            binding = entry.find(f"{{{OPENDOPE_XPATHS_NS}}}dataBinding")
            if binding is None:
                continue
            xpath = binding.get("xpath") or ""
            store_item_id = binding.get("storeItemID") or ""
            mappings = binding.get("prefixMappings") or ""
            value, part_name = _value_for(collection, store_item_id, xpath, mappings, limit)
            out.append(
                BindingInfo(
                    xpath=xpath,
                    prefix_mappings=mappings,
                    store_item_id=store_item_id,
                    kind="RichText",
                    tag="",
                    title=entry.get("name") or "",
                    value=value,
                    part_name=part_name,
                    open_dope_id=entry.get("id") or "",
                )
            )
    return tuple(out)


# ---------------------------------------------------------------------------
# fill()
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class FillEntry:
    """One key of a ``fill()``, and what became of it."""

    #: The key the caller gave: an XPath, a ``w:tag`` or a ``w:alias``.
    key: str
    #: The XPath it was matched to, or ``""``.
    xpath: str = ""
    #: The value written, as a string.
    value: str = ""
    #: Why it was skipped; ``""`` for one that was set.
    reason: str = ""
    #: A stable code for the reason.
    code: str = ""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        out: dict[str, Any] = {"key": self.key, "xpath": self.xpath, "value": self.value}
        if self.reason:
            out["reason"] = self.reason
            out["code"] = self.code
        return out


@dataclass(frozen=True, slots=True)
class FillResult:
    """What a ``fill()`` set, what it skipped, and what the bindings then did."""

    #: The keys that reached a node.
    applied: tuple[FillEntry, ...] = ()
    #: The keys that did not, each with its reason.
    skipped: tuple[FillEntry, ...] = ()
    #: The result of the :func:`apply_bindings` that followed.
    bindings: BindingResult = field(default_factory=lambda: BindingResult("apply_bindings"))
    #: The XPaths of the data nodes a repeat's list created (CR-003 section 17.10).
    created: tuple[str, ...] = ()
    #: The XPaths of the surplus data nodes it removed.
    removed: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: what a tool result returns."""
        return {
            "applied": [entry.to_dict() for entry in self.applied],
            "skipped": [entry.to_dict() for entry in self.skipped],
            "created": list(self.created),
            "removed": list(self.removed),
            "bindings": self.bindings.to_dict(),
        }

    def to_json(self, **options: Any) -> str:
        """:meth:`to_dict` as JSON."""
        return json.dumps(self.to_dict(), **options)

    def __repr__(self) -> str:
        """``<FillResult 4 set, 1 skipped, 20 bound>``."""
        return (
            f"<FillResult {len(self.applied)} set, {len(self.skipped)} skipped, "
            f"{self.bindings.updated} bound>"
        )


def fill_template(package: Any, data: dict[str, Any] | str) -> FillResult:
    """Set the nodes `data` names and apply the bindings. See the module docstring.

    Args:
        data: a mapping of XPath, ``w:tag`` or ``w:alias`` to value, or the XML
            of a whole custom XML part as a string. A value that is a **list**
            fills a repeat: the key is the repeating section's XPath (or the
            parent of its items), and each entry is a mapping of the fields
            ``describe().repeats`` reports --- or a plain value for a repeat
            whose node has no children.

    Raises:
        BindingError: a string was given and the document has no single
            non-built-in part to replace; or a list was given for a key that is
            not a repeat, or an entry names a field the template node has not.
    """
    from docx4j_py.model.content.reports import recording_on

    with recording_on(package, "fill") as change:
        return _fill(package, data, change)


def _fill(package: Any, data: dict[str, Any] | str, change: Any) -> FillResult:
    """The body of :func:`fill_template`, inside its **one** report.

    One report for the whole call, as ``insert_markdown`` has (CR-003 section
    17.9): a fill that writes a dozen nodes and then applies twenty bindings
    leaves one ``ChangeReport``, not thirty-three, because every call inside it
    finds one already open and is a no-op.
    """
    collection = custom_xml_parts_of(package)
    if isinstance(data, str):
        target = _single_part(collection)
        target.set_xml(data)
        bindings = apply_bindings(package)
        change.part(target.part)
        change.text(after=f"the whole of {target.part.part_name}")
        return FillResult(
            applied=(FillEntry(key=str(target.part.part_name), xpath="/", value="<xml>"),),
            bindings=bindings,
        )

    skeleton = describe_template(package)
    applied: list[FillEntry] = []
    skipped: list[FillEntry] = []
    created: list[str] = []
    removed: list[str] = []
    for key, value in data.items():
        if isinstance(value, (list, tuple)):
            entry, made, gone = _fill_repeat(collection, skeleton, key, list(value), change)
            applied.append(entry)
            created.extend(made)
            removed.extend(gone)
            continue
        text = "" if value is None else str(value)
        matches = _matches(skeleton, key)
        if not matches:
            skipped.append(
                FillEntry(
                    key=key,
                    value=text,
                    reason="no binding has this XPath, tag or title",
                    code="no_binding",
                )
            )
            continue
        wrote = False
        for binding in matches:
            part = collection.get_item(binding.store_item_id)
            if part is None:
                skipped.append(
                    FillEntry(
                        key=key,
                        xpath=binding.xpath,
                        value=text,
                        reason=f"no custom XML part with store item id {binding.store_item_id}",
                        code="no_part",
                    )
                )
                continue
            node = part.select_single_node(binding.xpath, binding.prefix_mappings)
            if node is None:
                skipped.append(
                    FillEntry(
                        key=key,
                        xpath=binding.xpath,
                        value=text,
                        reason=f"{binding.xpath!r} selects nothing in {part.part.part_name}",
                        code="no_match",
                    )
                )
                continue
            node.text = text
            wrote = True
            applied.append(FillEntry(key=key, xpath=binding.xpath, value=text))
            change.part(part.part)
            break
        if not wrote and not matches:  # pragma: no cover - guarded above
            continue
    bindings = apply_bindings(package)
    for entry in bindings.applied:
        change.touched(entry.address)
    for body in bound_bodies(package):
        if any(entry.address for entry in bindings.applied):
            change.part(getattr(body, "part", None))
    counts = [f"{len(applied)} set", f"{len(skipped)} skipped"]
    if created:
        counts.append(f"{len(created)} node{'' if len(created) == 1 else 's'} created")
    if removed:
        counts.append(f"{len(removed)} node{'' if len(removed) == 1 else 's'} removed")
    counts.append(f"{bindings.updated} bindings applied")
    change.text(after=", ".join(counts))
    for entry in skipped:
        change.warn(f"{entry.key}: {entry.reason}")
    return FillResult(tuple(applied), tuple(skipped), bindings, tuple(created), tuple(removed))


# ---------------------------------------------------------------------------
# a repeat's list of items (CR-003 section 17.10)
# ---------------------------------------------------------------------------


def _fill_repeat(
    collection: Any, skeleton: Skeleton, key: str, entries: list[Any], change: Any
) -> tuple[FillEntry, list[str], list[str]]:
    """One node per list entry, cloned from the template node the document has.

    The document's **one** ``w15:repeatingSectionItem`` is left exactly as it
    is: Word clones it to the node set when it opens the file, and a second
    template item would give a reader twice as many rows as the data has
    (CR-003 section 17.10).
    """
    repeat = _repeat_for(skeleton, key)
    if repeat is None:
        known = ", ".join(info.xpath for info in skeleton.repeats) or "none"
        raise BindingError(
            f"a list fills a repeating section, and {key!r} is not one this document binds",
            code="binding.not_a_repeat",
            hint=(
                f"the repeats are: {known}; pass a list under one of those XPaths "
                "(or the parent of its items), and a plain value under any other key"
            ),
        )
    part = collection.get_item(repeat.store_item_id) if repeat.store_item_id else None
    template = _repeat_node(collection, repeat.xpath, repeat.prefix_mappings, repeat.store_item_id)
    if part is None or template is None:
        raise BindingError(
            f"{repeat.xpath!r} selects no node in the data, so there is nothing to repeat",
            code="binding.no_repeat_node",
            hint="fill the part with fill(xml) first, or give the data one item to clone",
        )
    if not entries:
        raise BindingError(
            f"an empty list would leave {repeat.xpath!r} with no node to repeat from",
            code="binding.empty_repeat",
            hint="pass at least one entry; the document keeps one template item either way",
        )
    for entry in entries:
        if isinstance(entry, dict):
            _check_fields(part, template, repeat, entry)
        elif repeat.fields:
            raise BindingError(
                f"an entry of {repeat.xpath!r} is a mapping of its fields, "
                f"not {type(entry).__name__}",
                code="binding.entry_not_a_mapping",
                hint=f"one entry takes: {', '.join(repeat.fields)}",
            )

    parent = template.getparent()
    if parent is None:  # pragma: no cover - a repeat's node is never the root
        raise BindingError(
            f"{repeat.xpath!r} is the document element, which cannot repeat",
            code="binding.no_repeat_node",
            hint="bind the repeating section to a child of the root",
        )
    nodes = [child for child in parent if child.tag == template.tag]
    created: list[str] = []
    removed: list[str] = []
    while len(nodes) < len(entries):
        last = nodes[-1]
        clone = copy.deepcopy(template)
        # the last node's tail is the whitespace before the closing tag; the
        # template's is the whitespace between two items, so they swap and the
        # data part stays indented as its author left it
        clone.tail, last.tail = last.tail, template.tail
        parent.insert(parent.index(last) + 1, clone)
        nodes.append(clone)
        created.append(canonical_xpath_of(clone).xpath)
    for surplus in nodes[len(entries) :]:
        removed.append(canonical_xpath_of(surplus).xpath)
        parent.remove(surplus)
    del nodes[len(entries) :]
    for element, entry in zip(nodes, entries, strict=True):
        _write_entry(part, element, repeat, entry)
    part.touch()
    change.part(part.part)
    for xpath in created:
        change.touched(xpath)
    return (
        FillEntry(
            key=key,
            xpath=repeat.xpath,
            value=f"{len(entries)} item{'' if len(entries) == 1 else 's'}",
        ),
        created,
        removed,
    )


def _repeat_for(skeleton: Skeleton, key: str) -> RepeatInfo | None:
    """The repeat a ``fill()`` key names: its own XPath, or the parent of its items."""
    for info in skeleton.repeats:
        if key == info.xpath or key == info.xpath.rpartition("/")[0]:
            return info
    for binding in skeleton.bindings:
        if binding.kind != "RepeatingSection":
            continue
        if key in (binding.tag, binding.title) and key:
            for info in skeleton.repeats:
                if info.xpath == binding.xpath:
                    return info
    return None


def _targets(part: Any, element: Any, repeat: RepeatInfo, field_name: str) -> list[CustomXmlNode]:
    """The nodes one entry field names: a child element, or a relative XPath."""
    if _PLAIN_NAME.match(field_name):
        return [
            CustomXmlNode(child, part)
            for child in element
            if isinstance(child.tag, str) and _local(child.tag) == field_name
        ]
    return part.select_from(element, field_name, repeat.prefix_mappings)


def _check_fields(part: Any, template: Any, repeat: RepeatInfo, entry: dict[str, Any]) -> None:
    """Refuse an entry that names something the template node has not."""
    unknown = [name for name in entry if not _targets(part, template, repeat, str(name))]
    if not unknown:
        return
    fields = ", ".join(repeat.fields) or "none; this repeat takes plain values"
    raise BindingError(
        f"{repeat.xpath!r} has no {', '.join(repr(str(name)) for name in unknown)} "
        "to fill in a repeat entry",
        code="binding.unknown_field",
        hint=f"the fields of one entry are: {fields}",
    )


def _write_entry(part: Any, element: Any, repeat: RepeatInfo, entry: Any) -> None:
    """Set one node from one list entry: a mapping of fields, or a plain value."""
    if not isinstance(entry, dict):
        CustomXmlNode(element, part).text = "" if entry is None else str(entry)
        return
    for name, value in entry.items():
        for node in _targets(part, element, repeat, str(name)):
            node.text = "" if value is None else str(value)


def _matches(skeleton: Skeleton, key: str) -> list[BindingInfo]:
    """The bindings a ``fill()`` key names: by XPath, then by tag, then by title."""
    for attribute in ("xpath", "tag", "title"):
        found = [
            binding for binding in skeleton.bindings if getattr(binding, attribute) == key and key
        ]
        if found:
            return found
    return []


def _single_part(collection: Any) -> CustomXmlPart:
    """The one part a string ``fill()`` replaces."""
    candidates = [part for part in collection.items if not part.built_in]
    if len(candidates) == 1:
        return candidates[0]
    raise BindingError(
        f"this document has {len(candidates)} custom XML parts of its own, "
        "so fill(xml) cannot tell which one to replace",
        code="binding.ambiguous_part",
        hint="call pkg.custom_xml_parts.get_item(id).set_xml(xml) for the one you mean",
    )
