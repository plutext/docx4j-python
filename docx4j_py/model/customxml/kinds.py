"""The typed content controls: what Office JS hangs off a ``ContentControl``.

CR-003 section 3.7. One view per ``Word.ContentControlType`` that has members of
its own, each over the element in ``w:sdtPr`` that types the control:

============================  ==================================================
``CheckboxContentControl``    ``w14:checkbox``, its ``w14:checked`` and the two
                              ``w14:checkedState`` / ``uncheckedState`` glyphs
``DatePickerContentControl``  ``w:date``: ``w:dateFormat``, ``w:lid``,
                              ``w:calendar``, ``w:storeMappedDataAs``,
                              ``w:fullDate``
``DropDownListContentControl`` ``w:dropDownList`` and its ``w:listItem``\\ s
``ComboBoxContentControl``    ``w:comboBox``, the same shape
``PictureContentControl``     ``w:picture``, and the picture the control holds
``RepeatingSectionContentControl`` ``w15:repeatingSection``: the items, the
                              section title, whether Word offers the ``+``
``GroupContentControl``       ``w:group``; no members of its own, as Office JS
============================  ==================================================

Each is None on a control of another kind, which is how a caller asks "what is
this?" without a type test: ``control.checkbox_content_control`` or nothing.
"""

from __future__ import annotations

import datetime
import re
from typing import TYPE_CHECKING, Any, Literal

from docx4j_py.model.content.errors import BindingError
from docx4j_py.wml import el
from docx4j_py.wml.sdt import CHECKBOX_FONT, W14_NS, W15_NS, W_NS, sdt_property

if TYPE_CHECKING:  # pragma: no cover
    from docx4j_py.model.content.controls import ContentControl

__all__ = [
    "CHECKED_SYMBOL",
    "UNCHECKED_SYMBOL",
    "CheckboxContentControl",
    "ComboBoxContentControl",
    "ContentControlListItem",
    "DatePickerContentControl",
    "DropDownListContentControl",
    "GroupContentControl",
    "ListContentControl",
    "PictureContentControl",
    "RepeatingSectionContentControl",
    "format_date",
]

#: What Word shows for a checked box, and for an unchecked one.
CHECKED_SYMBOL = "☒"
UNCHECKED_SYMBOL = "☐"

#: ``w:storeMappedDataAs``: how a date picker stores its bound value.
DateStorageFormat = Literal["dateTime", "date", "text"]


def _value_of(holder: Any, default: Any = None) -> Any:
    """``@w:val``, with an enum member's value rather than the member."""
    value = getattr(holder, "val", None)
    if value is None:
        return default
    return getattr(value, "value", value)


class _KindView:
    """The shared half: the control, and the ``w:sdtPr`` child it is over."""

    __slots__ = ("control", "element")

    def __init__(self, control: ContentControl, element: Any) -> None:
        """Build the view over a control and the element that types it."""
        #: The control.
        self.control = control
        #: The ``w:sdtPr`` child: ``w14:checkbox``, ``w:date``, ….
        self.element = element

    def __eq__(self, other: object) -> bool:
        """Two views of the same element are equal."""
        return type(other) is type(self) and other.element is self.element

    def __hash__(self) -> int:
        """Hashes by the element."""
        return hash(id(self.element))

    def __repr__(self) -> str:
        """``<CheckboxContentControl checked>``."""
        return f"<{type(self).__name__}>"


# ---------------------------------------------------------------------------
# the checkbox
# ---------------------------------------------------------------------------


def _symbol_of(state: Any, fallback: str) -> str:
    """The glyph a ``w14:checkedState`` names: a hex code point, or a character."""
    value = _value_of(state)
    if not value:
        return fallback
    text = str(value)
    if len(text) in (4, 5) and all(character in "0123456789abcdefABCDEF" for character in text):
        try:
            return chr(int(text, 16))
        except ValueError:  # pragma: no cover - the length test already guards this
            return fallback
    return text


def _code_of(symbol: str) -> str:
    """The four-digit hex code point Word stores a glyph as."""
    return f"{ord(symbol[0]):04X}" if symbol else "2610"


class CheckboxContentControl(_KindView):
    """A ``w14:checkbox`` control. Office JS ``Word.CheckboxContentControl``."""

    __slots__ = ()

    @property
    def is_checked(self) -> bool:
        """Whether the box is ticked. Office JS ``isChecked``."""
        value = _value_of(getattr(self.element, "checked", None))
        return str(value).lower() in ("1", "true", "on")

    @is_checked.setter
    def is_checked(self, value: bool) -> None:
        """Tick or untick the box, **and rewrite the glyph the control shows**.

        Word draws the box from the content, not from ``w14:checked``, so the two
        are written together --- which is what makes a document Word opens show
        the state this set.
        """
        from docx4j_py.w14 import el as w14_el

        checked = getattr(self.element, "checked", None)
        if checked is None:
            checked = w14_el.checked()
            self.element.checked = checked
        checked.val = "1" if value else "0"
        self.control.set_checkbox_glyph(
            self.checked_symbol if value else self.unchecked_symbol, self.font
        )

    @property
    def checked_symbol(self) -> str:
        """The glyph a ticked box shows. Extension (``w14:checkedState``)."""
        return _symbol_of(getattr(self.element, "checked_state", None), CHECKED_SYMBOL)

    @checked_symbol.setter
    def checked_symbol(self, symbol: str) -> None:
        """Set the glyph, as its code point, and the font if it has none."""
        from docx4j_py.w14 import el as w14_el

        state = getattr(self.element, "checked_state", None)
        if state is None:
            state = w14_el.checkedState()
            self.element.checked_state = state
        state.val = _code_of(symbol)
        state.font = state.font or CHECKBOX_FONT

    @property
    def unchecked_symbol(self) -> str:
        """The glyph an empty box shows. Extension (``w14:uncheckedState``)."""
        return _symbol_of(getattr(self.element, "unchecked_state", None), UNCHECKED_SYMBOL)

    @unchecked_symbol.setter
    def unchecked_symbol(self, symbol: str) -> None:
        """Set the glyph, as its code point, and the font if it has none."""
        from docx4j_py.w14 import el as w14_el

        state = getattr(self.element, "unchecked_state", None)
        if state is None:
            state = w14_el.uncheckedState()
            self.element.unchecked_state = state
        state.val = _code_of(symbol)
        state.font = state.font or CHECKBOX_FONT

    @property
    def font(self) -> str:
        """The font the glyphs are drawn in; ``"MS Gothic"`` by default. Extension."""
        for name in ("checked_state", "unchecked_state"):
            state = getattr(self.element, name, None)
            if state is not None and getattr(state, "font", None):
                return str(state.font)
        return CHECKBOX_FONT

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary."""
        return {
            "kind": "CheckBox",
            "is_checked": self.is_checked,
            "checked_symbol": self.checked_symbol,
            "unchecked_symbol": self.unchecked_symbol,
            "font": self.font,
        }


# ---------------------------------------------------------------------------
# the date picker
# ---------------------------------------------------------------------------

#: Month and weekday names. English only, and a ``w:lid`` of another language
#: falls back to them: no locale dependency was taken (CR-003 section 17), and
#: ``en-AU``, ``en-GB`` and ``en-US`` --- which is what Word writes in practice
#: --- name the months the same way.
_MONTHS = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)
_DAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
#: The .NET date-format tokens Word's ``w:dateFormat`` uses.
_TOKEN = re.compile(r"d{1,4}|M{1,4}|y{2,4}|H{1,2}|h{1,2}|m{1,2}|s{1,2}|tt|'[^']*'|\"[^\"]*\"")


def format_date(pattern: str, when: datetime.datetime, locale: str = "en-US") -> str:
    """A date in Word's ``w:dateFormat`` pattern (.NET's, not ``strftime``'s).

    ``d MMMM yyyy`` gives ``29 January 2015``, which is what
    ``samples/invoice2013.docx`` shows. Text in single or double quotes is
    literal. `locale` is accepted and recorded but only English names are
    offered (CR-003 section 17).
    """
    del locale  # see the module note: English names, no locale dependency

    def replace(match: re.Match[str]) -> str:
        token = match.group(0)
        if token[:1] in ("'", '"'):
            return token[1:-1]
        hour12 = when.hour % 12 or 12
        return {
            "d": str(when.day),
            "dd": f"{when.day:02d}",
            "ddd": _DAYS[when.weekday()][:3],
            "dddd": _DAYS[when.weekday()],
            "M": str(when.month),
            "MM": f"{when.month:02d}",
            "MMM": _MONTHS[when.month - 1][:3],
            "MMMM": _MONTHS[when.month - 1],
            "yy": f"{when.year % 100:02d}",
            "yyy": str(when.year),
            "yyyy": str(when.year),
            "H": str(when.hour),
            "HH": f"{when.hour:02d}",
            "h": str(hour12),
            "hh": f"{hour12:02d}",
            "m": str(when.minute),
            "mm": f"{when.minute:02d}",
            "s": str(when.second),
            "ss": f"{when.second:02d}",
            "tt": "AM" if when.hour < 12 else "PM",
        }.get(token, token)

    return _TOKEN.sub(replace, pattern)


class DatePickerContentControl(_KindView):
    """A ``w:date`` control. Office JS ``Word.DatePickerContentControl``."""

    __slots__ = ()

    @property
    def date_display_format(self) -> str:
        """``w:dateFormat``, a .NET pattern such as ``d MMMM yyyy``. Office JS."""
        return str(_value_of(getattr(self.element, "date_format", None), "") or "")

    @date_display_format.setter
    def date_display_format(self, value: str) -> None:
        """Set the pattern; ``""`` removes the element."""
        if not value:
            self.element.date_format = None
            return
        self.element.date_format = el.dateFormat(val=value)

    #: docx4j and the OOXML element call it ``dateFormat``; Office JS calls it
    #: ``dateDisplayFormat``. Both names read and write the same element.
    date_format = date_display_format

    @property
    def date_display_locale(self) -> str:
        """``w:lid``, the language the date is written in (``en-AU``). Office JS."""
        return str(_value_of(getattr(self.element, "lid", None), "") or "")

    @date_display_locale.setter
    def date_display_locale(self, value: str) -> None:
        """Set the language; ``""`` removes the element."""
        self.element.lid = el.lid(val=value) if value else None

    @property
    def date_calendar_type(self) -> str:
        """``w:calendar``; ``"gregorian"`` when there is none. Office JS."""
        return str(_value_of(getattr(self.element, "calendar", None), "gregorian") or "gregorian")

    @date_calendar_type.setter
    def date_calendar_type(self, value: str) -> None:
        """Set the calendar."""
        self.element.calendar = el.calendar(val=value)

    @property
    def date_storage_format(self) -> str:
        """``w:storeMappedDataAs``; ``"text"`` when there is none. Office JS."""
        return str(_value_of(getattr(self.element, "store_mapped_data_as", None), "text") or "text")

    @date_storage_format.setter
    def date_storage_format(self, value: DateStorageFormat | str) -> None:
        """Set how the bound value is stored."""
        self.element.store_mapped_data_as = el.storeMappedDataAs(val=str(value))

    @property
    def full_date(self) -> datetime.datetime | None:
        """``w:fullDate``, the date Word really holds, or None. Extension."""
        value = getattr(self.element, "full_date", None)
        if value is None:
            return None
        try:
            return value.to_datetime()
        except Exception:  # noqa: BLE001 - a malformed w:fullDate is no date
            return None

    @full_date.setter
    def full_date(self, value: datetime.datetime | None) -> None:
        """Set ``w:fullDate``; None removes it."""
        if value is None:
            self.element.full_date = None
            return
        from docx4j_xsdata.models.datatype import XmlDateTime

        if value.tzinfo is None:
            value = value.replace(tzinfo=datetime.UTC)
        self.element.full_date = XmlDateTime.from_datetime(value.astimezone(datetime.UTC))

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary."""
        full = self.full_date
        return {
            "kind": "DatePicker",
            "date_display_format": self.date_display_format,
            "date_display_locale": self.date_display_locale,
            "date_calendar_type": self.date_calendar_type,
            "date_storage_format": self.date_storage_format,
            "full_date": full.isoformat() if full is not None else None,
        }


# ---------------------------------------------------------------------------
# the lists
# ---------------------------------------------------------------------------


class ContentControlListItem:
    """One ``w:listItem``. Office JS ``Word.ContentControlListItem``."""

    __slots__ = ("element", "items")

    def __init__(self, items: list, element: Any) -> None:
        """Build the view over an entry and the list holding it."""
        #: The live list of entries, so :meth:`delete` can find this one.
        self.items = items
        #: The ``w:listItem``.
        self.element = element

    @property
    def display_text(self) -> str:
        """``w:displayText``, what the drop-down shows. Office JS ``displayText``."""
        return str(getattr(self.element, "display_text", None) or "")

    @display_text.setter
    def display_text(self, value: str) -> None:
        """Set what the drop-down shows."""
        self.element.display_text = value

    @property
    def value(self) -> str:
        """``w:value``, what the binding stores. Office JS ``value``."""
        return str(getattr(self.element, "value", None) or "")

    @value.setter
    def value(self, value: str) -> None:
        """Set what the binding stores."""
        self.element.value = value

    @property
    def index(self) -> int:
        """This entry's position in the list, or -1. Office JS ``index``."""
        for index, item in enumerate(self.items):
            if item is self.element:
                return index
        return -1

    def delete(self) -> None:
        """Remove this entry. Office JS ``delete``."""
        at = self.index
        if at >= 0:
            del self.items[at]

    def to_dict(self) -> dict[str, Any]:
        """``{"display_text": …, "value": …, "index": …}``."""
        return {"display_text": self.display_text, "value": self.value, "index": self.index}

    def __eq__(self, other: object) -> bool:
        """Two views of the same entry are equal."""
        return isinstance(other, ContentControlListItem) and other.element is self.element

    def __hash__(self) -> int:
        """Hashes by the entry."""
        return hash(id(self.element))

    def __repr__(self) -> str:
        """``<ContentControlListItem 'Apples' = 'apples'>``."""
        return f"<ContentControlListItem {self.display_text!r} = {self.value!r}>"


class ListContentControl(_KindView):
    """The shared shape of a drop-down and a combo box. Office JS ``ListContentControl``."""

    __slots__ = ()

    @property
    def _entries(self) -> list:
        from docx4j_py.child import ChildList

        items = getattr(self.element, "list_item", None)
        if items is None:
            items = ChildList([], owner=self.element)
            self.element.list_item = items
        return items

    @property
    def list_items(self) -> list[ContentControlListItem]:
        """Every entry, in order. Office JS ``listItems``."""
        entries = self._entries
        return [ContentControlListItem(entries, item) for item in entries]

    def add_list_item(
        self, display_text: str, value: str | None = None, index: int | None = None
    ) -> ContentControlListItem:
        """Add an entry. Office JS ``addListItem``; `value` defaults to `display_text`."""
        entries = self._entries
        item = el.listItem(display_text=display_text, value=display_text if value is None else value)
        if index is None or index >= len(entries):
            entries.append(item)
        else:
            entries.insert(max(0, index), item)
        return ContentControlListItem(entries, item)

    def delete_all_list_items(self) -> None:
        """Remove every entry. Office JS ``deleteAllListItems``."""
        del self._entries[:]

    @property
    def last_value(self) -> str:
        """``w:lastValue``, what Word last showed. Extension."""
        return str(getattr(self.element, "last_value", None) or "")

    @last_value.setter
    def last_value(self, value: str) -> None:
        """Record what was last shown."""
        self.element.last_value = value

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary."""
        return {
            "kind": type(self).__name__.replace("ContentControl", ""),
            "list_items": [item.to_dict() for item in self.list_items],
            "last_value": self.last_value,
        }


class DropDownListContentControl(ListContentControl):
    """A ``w:dropDownList``. Office JS ``Word.DropDownListContentControl``."""

    __slots__ = ()


class ComboBoxContentControl(ListContentControl):
    """A ``w:comboBox``. Office JS ``Word.ComboBoxContentControl``."""

    __slots__ = ()


# ---------------------------------------------------------------------------
# the picture, the repeating section and the group
# ---------------------------------------------------------------------------


class PictureContentControl(_KindView):
    """A ``w:picture`` control. Office JS ``Word.PictureContentControl``."""

    __slots__ = ()

    @property
    def inline_picture(self) -> Any:
        """The picture the control holds, or None. Office JS ``inlinePicture``."""
        for paragraph in self.control.paragraphs:
            pictures = paragraph.inline_pictures
            if pictures:
                return pictures[0]
        return None

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary."""
        picture = self.inline_picture
        return {"kind": "Picture", "has_picture": picture is not None}


class RepeatingSectionContentControl(_KindView):
    """A ``w15:repeatingSection``. Office JS ``Word.RepeatingSectionContentControl``."""

    __slots__ = ()

    @property
    def items(self) -> list[ContentControl]:
        """The ``w15:repeatingSectionItem`` controls inside. Office JS ``items``."""
        return [
            control
            for control in self.control.content_controls
            if control.type == "RepeatingSectionItem"
        ]

    @property
    def section_title(self) -> str:
        """``w15:sectionTitle``, what Word's ``+`` tooltip says. Office JS."""
        return str(_value_of(getattr(self.element, "section_title", None), "") or "")

    @section_title.setter
    def section_title(self, value: str) -> None:
        """Set the title; ``""`` removes the element."""
        from docx4j_py.w15 import el as w15_el

        self.element.section_title = w15_el.sectionTitle(val=value) if value else None

    @property
    def allow_insert_delete_section(self) -> bool:
        """Whether Word offers the ``+``. Office JS ``allowInsertDeleteSection``.

        The inverse of ``w15:doNotAllowInsertDeleteSection``, which is absent by
        default and therefore True.
        """
        flag = getattr(self.element, "do_not_allow_insert_delete_section", None)
        if flag is None:
            return True
        value = _value_of(flag, True)
        return str(value).lower() in ("0", "false")

    @allow_insert_delete_section.setter
    def allow_insert_delete_section(self, value: bool) -> None:
        """Allow or forbid inserting and deleting sections."""
        from docx4j_py.w15 import el as w15_el

        if value:
            self.element.do_not_allow_insert_delete_section = None
        else:
            self.element.do_not_allow_insert_delete_section = w15_el.doNotAllowInsertDeleteSection(
                val="1"
            )

    def insert_item_after(self, index: int) -> ContentControl:
        """Copy an item and put the copy after it, as Word's ``+`` does.

        Extension in name (Office JS has no such method), and a structural deep
        copy: the ``w:sdt`` of the item at `index` is copied, the copy's ``w:id``
        is re-allocated, and it goes in straight after the original.

        Raises:
            BindingError: there is no item at `index`.
        """
        items = self.items
        if not 0 <= index < len(items):
            raise BindingError(
                f"this repeating section has {len(items)} item(s); there is none at {index}",
                code="binding.no_such_item",
                hint="index from 0; read repeating_section_content_control.items first",
            )
        return items[index].insert_copy_after()

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary."""
        return {
            "kind": "RepeatingSection",
            "items": len(self.items),
            "section_title": self.section_title,
            "allow_insert_delete_section": self.allow_insert_delete_section,
        }


class GroupContentControl(_KindView):
    """A ``w:group``. Office JS ``Word.GroupContentControl``: no members of its own."""

    __slots__ = ()

    def to_dict(self) -> dict[str, Any]:
        """A JSON-ready summary."""
        return {"kind": "Group"}


#: Which ``w:sdtPr`` child each view is over, and the namespace it lives in.
KIND_VIEWS: dict[str, tuple[type[_KindView], str, str]] = {
    "CheckBox": (CheckboxContentControl, "checkbox", W14_NS),
    "DatePicker": (DatePickerContentControl, "date", W_NS),
    "DropDownList": (DropDownListContentControl, "dropDownList", W_NS),
    "ComboBox": (ComboBoxContentControl, "comboBox", W_NS),
    "Picture": (PictureContentControl, "picture", W_NS),
    "RepeatingSection": (RepeatingSectionContentControl, "repeatingSection", W15_NS),
    "Group": (GroupContentControl, "group", W_NS),
}


def kind_view(control: ContentControl, kind: str) -> Any:
    """The typed view of a control, or None when it is not that kind."""
    found = KIND_VIEWS.get(kind)
    if found is None or control.type != kind:
        return None
    view_class, local_name, namespace = found
    element = sdt_property(control.sdt_pr, local_name, namespace)
    return view_class(control, element) if element is not None else None
