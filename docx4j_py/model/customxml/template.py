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
    document shows the new values without waiting for Word to open it.

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

import json
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
from docx4j_py.model.customxml.parts import custom_xml_parts_of
from docx4j_py.wml.sdt import KIND_BY_QNAME, W15_NS, W_NS

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.customxml.parts import CustomXmlPart

__all__ = [
    "MAX_VALUE_CHARS",
    "OPENDOPE_XPATHS_NS",
    "BindingInfo",
    "FillEntry",
    "FillResult",
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
    #: The XPaths of the repeating sections, in document order.
    repeats: tuple[str, ...] = ()
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
            "repeats": list(self.repeats),
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
        repeats=repeats,
        parts=parts,
        open_dope=bool(open_dope_bindings),
        warnings=tuple(warnings),
    )


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
) -> tuple[tuple[BindingInfo, ...], tuple[str, ...]]:
    """The bindings read through the content-API views, addresses and all."""
    bindings: list[BindingInfo] = []
    repeats: list[str] = []
    for control in controls_of(package):
        mapping = control.xml_mapping
        if control.type == "RepeatingSection" and mapping.is_mapped:
            repeats.append(mapping.xpath)
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
) -> tuple[tuple[BindingInfo, ...], tuple[str, ...]]:
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
    repeats: list[str] = []
    for pr in root.iter(f"{_W}sdtPr"):
        binding = _binding_element(pr)
        kind = _kind_of(pr)
        xpath = binding.get(f"{_W}xpath") if binding is not None else None
        if kind == "RepeatingSection" and xpath:
            repeats.append(xpath)
        if binding is None or not xpath:
            continue
        mappings = binding.get(f"{_W}prefixMappings") or ""
        store_item_id = binding.get(f"{_W}storeItemID") or ""
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

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: what a tool result returns."""
        return {
            "applied": [entry.to_dict() for entry in self.applied],
            "skipped": [entry.to_dict() for entry in self.skipped],
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
            of a whole custom XML part as a string.

    Raises:
        BindingError: a string was given and the document has no single
            non-built-in part to replace.
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
    for key, value in data.items():
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
    change.text(
        after=f"{len(applied)} set, {len(skipped)} skipped, {bindings.updated} bindings applied"
    )
    for entry in skipped:
        change.warn(f"{entry.key}: {entry.reason}")
    return FillResult(tuple(applied), tuple(skipped), bindings)


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
