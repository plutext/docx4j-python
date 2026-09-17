"""``apply_bindings()`` and ``update_from_content_controls()``: the data, both ways.

CR-003 sections 3.7 and 4. docx4j's ``BindingHandler.applyBindings`` --- what
Word does when it opens a document whose controls are bound --- and its reverse,
``UpdateXmlFromDocumentSurface``, with the shapes docx4j uses and the four
departures CR-003 section 4 makes rules:

* **containers are never bound**, in either direction: a repeating section, a
  repeating-section item, a group, a building-block gallery, anything holding a
  nested control, and a block control holding a table. docx4j would collapse a
  ``w15:repeatingSection`` to one run;
* **pictures and explicitly rich-text controls are counted as skipped**, not
  failed: replacing an ``a:blip`` from base64 and unpacking a flat OPC package
  are separate work;
* a **date** is formatted with ``w:dateFormat`` in the ``w:lid`` locale and
  ``w:fullDate`` is set; a **checkbox** gets ``w14:checked`` *and* the glyph run
  from ``w14:checkedState`` / ``uncheckedState``; a **list** shows the
  ``w:displayText`` whose ``w:value`` matches. The reverse writes ``true`` /
  ``false``, the entry's ``w:value`` and the stored ``w:fullDate``, where docx4j
  writes the rendered glyph and skips dates altogether;
* the value goes into the control's **first paragraph** for a block, row or cell
  control, and ``w:placeholder`` is kept.

Everything else is docx4j's: the value is trimmed, the run properties come from
``w:sdtPr/w:rPr``, an empty value becomes the placeholder run and
``w:showingPlcHdr``, ``xml:space="preserve"`` is written only for a leading or
trailing space, and a multiline ``w:text`` control turns newlines into ``w:br``.
"""

from __future__ import annotations

import datetime
import json
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from docx4j_py.child import ChildList
from docx4j_py.model.customxml.kinds import format_date
from docx4j_py.wml import R, el, t
from docx4j_py.wml.sdt import W_NS, sdt_property

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.controls import ContentControl

__all__ = [
    "CONTAINER_KINDS",
    "PLACEHOLDER_STYLE",
    "PLACEHOLDER_TEXT",
    "BindingEntry",
    "BindingResult",
    "apply_bindings",
    "bound_bodies",
    "controls_of",
    "runs_for_value",
    "update_from_content_controls",
    "write_control_to_node",
]

#: docx4j's placeholder: the style and the words it writes for an empty value.
PLACEHOLDER_STYLE = "PlaceholderText"
PLACEHOLDER_TEXT = "Click here to enter text."

#: The kinds that hold other controls rather than a value of their own.
CONTAINER_KINDS: frozenset[str] = frozenset(
    {"RepeatingSection", "RepeatingSectionItem", "Group", "BuildingBlockGallery"}
)

#: What a date value has to look like before it is read as one.
_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2})?)?")

#: The strings Word reads as a ticked box.
_TRUE = frozenset({"true", "1", "on", "yes"})


# ---------------------------------------------------------------------------
# what a run reports
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class BindingEntry:
    """One binding, and what became of it. CR-003 section 3.1's frozen result."""

    #: The bound XPath.
    xpath: str
    #: The ``w:storeItemID`` of the part it names.
    store_item_id: str
    #: The Office JS ``Word.ContentControlType``.
    kind: str
    #: ``w:tag``, the machine-readable name.
    tag: str
    #: ``w:alias``, which Word's dialog calls the title.
    title: str
    #: The control's ordinal address.
    address: str
    #: The value written, or read, or ``""``.
    value: str = ""
    #: Why it was skipped; ``""`` for one that was applied.
    reason: str = ""
    #: A stable code for the reason: ``"container"``, ``"picture"``, ….
    code: str = ""

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view."""
        out: dict[str, Any] = {
            "xpath": self.xpath,
            "store_item_id": self.store_item_id,
            "kind": self.kind,
            "tag": self.tag,
            "title": self.title,
            "address": self.address,
            "value": self.value,
        }
        if self.reason:
            out["reason"] = self.reason
            out["code"] = self.code
        return out


@dataclass(frozen=True, slots=True)
class BindingResult:
    """What one pass over the bindings did. CR-003 section 3.1."""

    #: ``"apply_bindings"`` or ``"update_from_content_controls"``.
    operation: str
    #: The bindings that were written.
    applied: tuple[BindingEntry, ...] = ()
    #: The bindings that were not, each with its reason.
    skipped: tuple[BindingEntry, ...] = ()

    @property
    def bound(self) -> int:
        """How many controls carried a binding at all."""
        return len(self.applied) + len(self.skipped)

    @property
    def updated(self) -> int:
        """How many were written."""
        return len(self.applied)

    def to_dict(self) -> dict[str, Any]:
        """The JSON-ready view: what a tool result returns."""
        return {
            "operation": self.operation,
            "bound": self.bound,
            "updated": self.updated,
            "applied": [entry.to_dict() for entry in self.applied],
            "skipped": [entry.to_dict() for entry in self.skipped],
        }

    def to_json(self, **options: Any) -> str:
        """:meth:`to_dict` as JSON."""
        return json.dumps(self.to_dict(), **options)

    def __repr__(self) -> str:
        """``<BindingResult apply_bindings 17 of 20>``."""
        return f"<BindingResult {self.operation} {self.updated} of {self.bound}>"


# ---------------------------------------------------------------------------
# finding the controls
# ---------------------------------------------------------------------------


def bound_bodies(package: Any) -> list[Any]:
    """The bodies binding looks in: the main document part, the headers, the footers.

    docx4j's ``BindingHandler.applyBindings`` walks exactly these three.
    """
    from docx4j_py.model.content.body import body_of
    from docx4j_py.model.content.errors import ContentError

    parts: list[Any] = []
    main = getattr(package, "main_document_part", None)
    if main is not None:
        parts.append(main)
    for name in ("header_parts", "footer_parts"):
        getter = getattr(package, name, None)
        if getter is not None:
            parts.extend(getter())
    out: list[Any] = []
    for part in parts:
        # ``body_of(part)`` and not ``part.body``: a trial's ``TrialPart`` has
        # no ``body`` of its own and would hand the attribute to the **real**
        # part, so a dry run's ``apply_bindings`` would edit the real document
        # (CR-003 section 17).
        try:
            out.append(body_of(part))
        except ContentError:  # pragma: no cover - a part with no block content
            continue
    return out


def controls_of(package: Any) -> list[ContentControl]:
    """Every content control in the bound bodies, in document order."""
    out: list[ContentControl] = []
    for body in bound_bodies(package):
        out.extend(body.content_controls)
    return out


def is_container(control: ContentControl) -> bool:
    """Whether a control holds other controls rather than a value (CR-003 section 4)."""
    if control.type in CONTAINER_KINDS:
        return True
    if control.content_controls:
        return True
    return bool(control.form == "block" and control.tables)


# ---------------------------------------------------------------------------
# the runs a value becomes (docx4j ValueInserterPlainTextImpl)
# ---------------------------------------------------------------------------


def runs_for_value(value: str, run_properties: Any, multiline: bool) -> list[Any]:
    """The runs docx4j writes for a bound value.

    An empty value is the placeholder run; a multiline one becomes a run per
    line with a ``w:br`` run between them; ``xml:space="preserve"`` is written
    only when a line starts or ends with a space.
    """
    from docx4j_py.child import deep_copy

    def properties() -> Any:
        return deep_copy(run_properties) if run_properties is not None else None

    if value == "":
        rpr = properties() or el.rPr()
        rpr.r_style = el.rStyle(val=PLACEHOLDER_STYLE)
        return [R(r_pr=rpr, content=ChildList([t(PLACEHOLDER_TEXT)]))]

    if multiline:
        lines = [line for line in re.split(r"[\n\r\f]+", value) if line != ""]
    else:
        lines = [re.sub(r"[\n\r\f]+", "", value)]

    runs: list[Any] = []
    for index, line in enumerate(lines):
        if index:
            runs.append(R(r_pr=properties(), content=ChildList([el.br()])))
        text = t(line)
        if line.startswith(" ") or line.endswith(" "):
            text.space = "preserve"
        runs.append(R(r_pr=properties(), content=ChildList([text])))
    return runs


# ---------------------------------------------------------------------------
# the data into the controls
# ---------------------------------------------------------------------------


def _entry(control: ContentControl, value: str = "", reason: str = "", code: str = "") -> Any:
    mapping = control.xml_mapping
    return BindingEntry(
        xpath=mapping.xpath,
        store_item_id=mapping.store_item_id,
        kind=control.type,
        tag=control.tag,
        title=control.title,
        address=control.address,
        value=value,
        reason=reason,
        code=code,
    )


def _skip_reason(control: ContentControl) -> tuple[str, str] | None:
    """Why this bound control is not written, or None."""
    if control.type == "Picture":
        return "a picture binding replaces the image, which this phase does not do", "picture"
    if sdt_property(control.sdt_pr, "richText", W_NS) is not None:
        return "an explicitly rich-text binding carries a package, not a value", "rich_text"
    if is_container(control):
        return "a container is never bound; its items are", "container"
    return None


def date_of(value: str) -> datetime.datetime | None:
    """A bound value read as a date, or None. Word's stored form, ``Z`` and all."""
    if not _DATE.match(value):
        return None
    text = value.rstrip("Z").replace(" ", "T")
    try:
        return datetime.datetime.fromisoformat(text)
    except ValueError:
        return None


def iso_of(when: datetime.datetime) -> str:
    """Word's stored date form, ``2015-01-29T00:00:00``."""
    return when.replace(tzinfo=None, microsecond=0).isoformat()


def apply_binding_to(control: ContentControl) -> tuple[bool, str, str, str]:
    """Push one control's data into it.

    Returns:
        ``(applied, value, reason, code)``.
    """
    mapping = control.xml_mapping
    node = mapping.custom_xml_node
    if node is None:
        part = mapping.custom_xml_part
        if part is None:
            return (
                False,
                "",
                f"no custom XML part with store item id {mapping.store_item_id}",
                "no_part",
            )
        return False, "", f"{mapping.xpath!r} selects nothing in {part.part.part_name}", "no_match"

    skip = _skip_reason(control)
    if skip is not None:
        return False, node.text.strip(), skip[0], skip[1]

    value = node.text.strip()

    if control.type == "CheckBox":
        checkbox = control.checkbox_content_control
        if checkbox is None:
            return False, value, "the control has no w14:checkbox to set", "no_checkbox"
        checkbox.is_checked = value.lower() in _TRUE
        return True, value, "", ""

    display = value
    if control.type in ("DropDownList", "ComboBox"):
        listing = control.drop_down_list_content_control or control.combo_box_content_control
        if listing is not None:
            match = next((item for item in listing.list_items if item.value == value), None)
            if match is not None:
                display = match.display_text
            listing.last_value = value
    elif control.type == "DatePicker":
        date_view = control.date_picker_content_control
        parsed = date_of(value)
        if date_view is not None and parsed is not None:
            date_view.full_date = parsed
            if date_view.date_display_format:
                display = format_date(
                    date_view.date_display_format,
                    parsed,
                    date_view.date_display_locale or "en-US",
                )

    control.set_bound_content(
        runs_for_value(display, control.run_properties, control.is_multi_line),
        display == "",
    )
    return True, display, "", ""


def apply_bindings(package: Any) -> BindingResult:
    """Push the custom XML into every bound control. docx4j ``BindingHandler``."""
    applied: list[BindingEntry] = []
    skipped: list[BindingEntry] = []
    for control in controls_of(package):
        if not control.xml_mapping.is_mapped:
            continue
        done, value, reason, code = apply_binding_to(control)
        (applied if done else skipped).append(_entry(control, value, reason, code))
    return BindingResult("apply_bindings", tuple(applied), tuple(skipped))


# ---------------------------------------------------------------------------
# the controls back into the data
# ---------------------------------------------------------------------------


def update_from_control(control: ContentControl) -> tuple[bool, str, str, str]:
    """Write one control's value back into its node.

    Returns:
        ``(written, value, reason, code)``.
    """
    mapping = control.xml_mapping
    node = mapping.custom_xml_node
    if node is None:
        part = mapping.custom_xml_part
        if part is None:
            return (
                False,
                "",
                f"no custom XML part with store item id {mapping.store_item_id}",
                "no_part",
            )
        return False, "", f"{mapping.xpath!r} selects nothing in {part.part.part_name}", "no_match"

    skip = _skip_reason(control)
    if skip is not None:
        return False, "", skip[0], skip[1]
    if control.is_showing_placeholder:
        return False, "", "a control showing its placeholder holds no value", "placeholder"

    value = control.text
    if control.type == "CheckBox":
        checkbox = control.checkbox_content_control
        if checkbox is None:
            return False, "", "the control has no w14:checkbox to read", "no_checkbox"
        value = "true" if checkbox.is_checked else "false"
    elif control.type in ("DropDownList", "ComboBox"):
        listing = control.drop_down_list_content_control or control.combo_box_content_control
        if listing is not None:
            match = next((item for item in listing.list_items if item.display_text == value), None)
            if match is not None:
                value = match.value
    elif control.type == "DatePicker":
        date_view = control.date_picker_content_control
        full = date_view.full_date if date_view is not None else None
        if full is not None:
            value = iso_of(full)

    if node.text == value:
        return False, value, "the node already holds this value", "unchanged"
    node.text = value
    return True, value, "", ""


def update_from_content_controls(package: Any) -> BindingResult:
    """Write every bound control back into the custom XML. docx4j's reverse pass."""
    applied: list[BindingEntry] = []
    skipped: list[BindingEntry] = []
    for control in controls_of(package):
        if not control.xml_mapping.is_mapped:
            continue
        done, value, reason, code = update_from_control(control)
        (applied if done else skipped).append(_entry(control, value, reason, code))
    return BindingResult("update_from_content_controls", tuple(applied), tuple(skipped))


def write_control_to_node(control: ContentControl) -> bool:
    """One control's write-through, for :meth:`ContentControl.insert_text`.

    Silent about everything it skips: a control that is not bound, whose part is
    gone or whose XPath no longer resolves is simply not written, because an
    ``insert_text`` is a text edit and not a binding operation.
    """
    if not control.xml_mapping.is_mapped:
        return False
    try:
        done, _value, _reason, _code = update_from_control(control)
    except Exception:  # noqa: BLE001 - an unusable binding must not fail a text edit
        return False
    return done
