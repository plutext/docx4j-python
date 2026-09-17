"""``document.styles``: python-docx's ``Styles`` and its style objects.

CR-003 section 3.6, Phase J. python-docx addresses styles by their **display
name** (``styles["Heading 1"]``) and writes their **id** into ``w:pStyle``; the
translation is CR-003 section 4's, and this module uses the content API's own
table for it (:func:`~docx4j_py.model.content.styles.display_name_of`,
:func:`~docx4j_py.model.content.styles.style_id_of`) rather than a second copy
of python-docx's ``BabelFish``.

Two things differ from python-docx and are recorded in section 20:

* a **built-in style the document does not define is defined** when a setter
  names it (CR-003 section 14.9), where python-docx raises --- because its own
  default template defines them all and this engine's ``create_package()``
  defines five;
* ``latent_styles`` is **not** offered: it is 371 ``w:lsdException`` entries
  behind an API nothing in the facade's tests or in the wild uses, and the
  typed ``w:latentStyles`` is one attribute away on ``styles.element``.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any

from docx4j_py.docx._base import recording
from docx4j_py.docx.enums import WD_STYLE_TYPE
from docx4j_py.model.content.styles import display_name_of, style_id_of
from docx4j_py.wml import Style, StyleNext
from docx4j_py.wml import Styles as StylesElement

__all__ = [
    "BaseStyle",
    "CharacterStyle",
    "ParagraphStyle",
    "Styles",
    "_NumberingStyle",
    "_TableStyle",
    "style_for",
    "style_id_for",
    "styles_of",
]


def styles_of(package: Any, *, create: bool = False) -> StylesElement | None:
    """The unmarshalled ``w:styles`` of a package, or None when there is no part.

    **This unmarshals ``/word/styles.xml``**, which marks it for
    re-marshalling: python-docx's style objects are over the part, so there is
    no reading one without it. Everything the facade can answer without the
    part --- ``paragraph.text``, ``document.paragraphs`` --- does.
    """
    part = getattr(package, "style_definitions_part", None) if package is not None else None
    if part is None:
        return None
    styles = part.contents
    if create and styles.style is None:
        from docx4j_py.child import ChildList

        styles.style = ChildList([], owner=styles)
    return styles


class BaseStyle:
    """One ``w:style``: what every kind of style object has."""

    __slots__ = ("_package", "_style")

    def __init__(self, style: Style, package: Any = None) -> None:
        """Wrap the ``w:style``."""
        self._style = style
        self._package = package

    def __eq__(self, other: object) -> bool:
        """Two views of one ``w:style`` are equal."""
        return isinstance(other, BaseStyle) and other._style is self._style

    def __hash__(self) -> int:
        """Hashes by the element's identity."""
        return hash(id(self._style))

    def __repr__(self) -> str:
        """``<docx4j_py.docx ParagraphStyle 'Heading 1'>``."""
        return f"<docx4j_py.docx {type(self).__name__} {self.name!r}>"

    @property
    def element(self) -> Style:
        """The typed ``w:style``."""
        return self._style

    @property
    def part(self) -> Any:
        """``/word/styles.xml``, or None when the document has no styles part.

        python-docx raises here, because ``StyleFactory`` builds a style object
        with no parent at all; the part is reachable in this engine, so it is
        answered.
        """
        return (
            getattr(self._package, "style_definitions_part", None)
            if self._package is not None
            else None
        )

    # -- identity ----------------------------------------------------------

    @property
    def name(self) -> str | None:
        """The style's display name (``"Heading 1"`` for ``w:name`` ``heading 1``)."""
        stored = getattr(getattr(self._style, "name", None), "val", None)
        style_id = self.style_id
        if stored is None and not style_id:
            return None
        return display_name_of(style_id or "", stored)

    @name.setter
    def name(self, value: str) -> None:
        with self._recording("docx.style.name"):
            from docx4j_py.wml import el

            self._style.name = el.name(val=value)
            self._style.name.parent = self._style

    @property
    def style_id(self) -> str:
        """``w:styleId``."""
        return self._style.style_id or ""

    @style_id.setter
    def style_id(self, value: str) -> None:
        with self._recording("docx.style.style_id"):
            self._style.style_id = value

    @property
    def type(self) -> WD_STYLE_TYPE:
        """The style's kind; ``PARAGRAPH`` when ``w:type`` is absent."""
        value = getattr(self._style, "type_value", None)
        value = getattr(value, "value", value)
        if value is None:
            return WD_STYLE_TYPE.PARAGRAPH
        try:
            return WD_STYLE_TYPE.from_xml(str(value))
        except ValueError:
            return WD_STYLE_TYPE.PARAGRAPH

    @property
    def builtin(self) -> bool:
        """True when ``w:customStyle`` is absent, as python-docx decides it."""
        return not bool(self._style.custom_style)

    # -- the on/off attributes --------------------------------------------

    def _toggle(self, field: str) -> bool:
        toggle = getattr(self._style, field, None)
        if toggle is None:
            return False
        value = getattr(toggle, "val", None)
        return True if value is None else bool(value)

    def _set_toggle(self, field: str, name: str, value: bool | None) -> None:
        from docx4j_py.wml import el

        with self._recording(f"docx.style.{field}"):
            if not value:
                setattr(self._style, field, None)
                return
            toggle = getattr(el, name)()
            setattr(self._style, field, toggle)
            toggle.parent = self._style

    @property
    def hidden(self) -> bool:
        """``w:semiHidden``: kept out of the style gallery."""
        return self._toggle("semi_hidden")

    @hidden.setter
    def hidden(self, value: bool | None) -> None:
        self._set_toggle("semi_hidden", "semiHidden", value)

    @property
    def locked(self) -> bool:
        """``w:locked``: not applicable while formatting protection is on."""
        return self._toggle("locked")

    @locked.setter
    def locked(self, value: bool | None) -> None:
        self._set_toggle("locked", "locked", value)

    @property
    def quick_style(self) -> bool:
        """``w:qFormat``: shown in the style gallery."""
        return self._toggle("q_format")

    @quick_style.setter
    def quick_style(self, value: bool | None) -> None:
        self._set_toggle("q_format", "qFormat", value)

    @property
    def unhide_when_used(self) -> bool:
        """``w:unhideWhenUsed``."""
        return self._toggle("unhide_when_used")

    @unhide_when_used.setter
    def unhide_when_used(self, value: bool | None) -> None:
        self._set_toggle("unhide_when_used", "unhideWhenUsed", value)

    @property
    def priority(self) -> int | None:
        """``w:uiPriority``: the sort key Word shows the style under."""
        value = getattr(getattr(self._style, "ui_priority", None), "val", None)
        return None if value is None else int(value)

    @priority.setter
    def priority(self, value: int | None) -> None:
        from docx4j_py.wml import el

        with self._recording("docx.style.priority"):
            if value is None:
                self._style.ui_priority = None
                return
            self._style.ui_priority = el.uiPriority(val=int(value))
            self._style.ui_priority.parent = self._style

    # -- the graph ---------------------------------------------------------

    def _sibling(self, style_id: str | None) -> BaseStyle | None:
        if not style_id:
            return None
        styles = getattr(self._style, "parent", None)
        for style in getattr(styles, "style", None) or ():
            if style.style_id == style_id:
                return style_factory(style, self._package)
        return None

    @property
    def base_style(self) -> BaseStyle | None:
        """The style this one is based on (``w:basedOn``), or None."""
        return self._sibling(getattr(getattr(self._style, "based_on", None), "val", None))

    @base_style.setter
    def base_style(self, style: BaseStyle | None) -> None:
        from docx4j_py.wml import el

        with self._recording("docx.style.base_style"):
            if style is None:
                self._style.based_on = None
                return
            self._style.based_on = el.basedOn(val=style.style_id)
            self._style.based_on.parent = self._style

    def delete(self) -> None:
        """Remove this style definition from ``/word/styles.xml``."""
        with self._recording("docx.style.delete"):
            styles = getattr(self._style, "parent", None)
            items = getattr(styles, "style", None)
            if items is not None:
                for index, style in enumerate(list(items)):
                    if style is self._style:
                        del items[index]
                        break

    # -- formatting --------------------------------------------------------

    @property
    def font(self) -> Any:
        """The style's :class:`~docx4j_py.docx.text.Font`, over its ``w:rPr``."""
        from docx4j_py.docx.text import Font

        return Font(self._style, None)

    def _recording(self, operation: str) -> Any:
        """A change report naming ``/word/styles.xml``."""
        from docx4j_py.model.content.reports import recording_on

        part = (
            getattr(self._package, "style_definitions_part", None)
            if self._package is not None
            else None
        )
        return recording_on(self._package, operation, *([part] if part is not None else []))


class CharacterStyle(BaseStyle):
    """A character style: ``w:style/@w:type="character"``."""

    __slots__ = ()


class ParagraphStyle(CharacterStyle):
    """A paragraph style: character formatting plus ``paragraph_format``."""

    __slots__ = ()

    @property
    def next_paragraph_style(self) -> ParagraphStyle:
        """``w:next`` as a style object; this style when there is none."""
        nxt = self._sibling(getattr(getattr(self._style, "next", None), "val", None))
        if nxt is None or nxt.type is not WD_STYLE_TYPE.PARAGRAPH:
            return self
        return nxt  # type: ignore[return-value]

    @next_paragraph_style.setter
    def next_paragraph_style(self, style: ParagraphStyle | None) -> None:

        with self._recording("docx.style.next_paragraph_style"):
            if style is None or style.style_id == self.style_id:
                self._style.next = None
                return
            self._style.next = StyleNext(val=style.style_id)
            self._style.next.parent = self._style

    @property
    def paragraph_format(self) -> Any:
        """The style's :class:`~docx4j_py.docx.text.ParagraphFormat`."""
        from docx4j_py.docx.text import ParagraphFormat

        return ParagraphFormat(self._style, None)


class _TableStyle(ParagraphStyle):
    """A table style: ``w:style/@w:type="table"``."""

    __slots__ = ()


class _NumberingStyle(BaseStyle):
    """A numbering style: ``w:style/@w:type="numbering"``. Nothing of its own."""

    __slots__ = ()


_BY_TYPE: dict[WD_STYLE_TYPE, type[BaseStyle]] = {
    WD_STYLE_TYPE.PARAGRAPH: ParagraphStyle,
    WD_STYLE_TYPE.CHARACTER: CharacterStyle,
    WD_STYLE_TYPE.TABLE: _TableStyle,
    WD_STYLE_TYPE.LIST: _NumberingStyle,
}


def style_factory(style: Style, package: Any = None) -> BaseStyle:
    """The right style object for a ``w:style``. python-docx's ``StyleFactory``."""
    value = getattr(style, "type_value", None)
    value = getattr(value, "value", value)
    kind = WD_STYLE_TYPE.PARAGRAPH
    if value is not None:
        try:
            kind = WD_STYLE_TYPE.from_xml(str(value))
        except ValueError:
            kind = WD_STYLE_TYPE.PARAGRAPH
    return _BY_TYPE[kind](style, package)


class Styles:
    """``document.styles``: the styles the document defines.

    Supports ``len()``, iteration and lookup by display name, as python-docx's
    does. Touching it **unmarshals ``/word/styles.xml``**.
    """

    __slots__ = ("_package",)

    def __init__(self, package: Any) -> None:
        """Wrap the package whose styles part this is over."""
        self._package = package

    def __repr__(self) -> str:
        """``<docx4j_py.docx Styles, 12 defined>``."""
        return f"<docx4j_py.docx Styles, {len(self)} defined>"

    @property
    def part(self) -> Any:
        """``/word/styles.xml``, or None when the document has no styles part."""
        return getattr(self._package, "style_definitions_part", None)

    @property
    def element(self) -> StylesElement:
        """The typed ``w:styles``."""
        styles = styles_of(self._package)
        if styles is None:
            raise ValueError("this document has no /word/styles.xml")
        return styles

    @property
    def _list(self) -> list[Style]:
        styles = styles_of(self._package)
        return list(getattr(styles, "style", None) or ()) if styles is not None else []

    def __len__(self) -> int:
        """How many ``w:style`` definitions there are."""
        return len(self._list)

    def __iter__(self) -> Iterator[BaseStyle]:
        """Each style object, in document order."""
        return (style_factory(style, self._package) for style in self._list)

    def __contains__(self, name: object) -> bool:
        """``"Heading 1" in document.styles``, by display name."""
        return self._find(str(name)) is not None

    def __getitem__(self, key: str) -> BaseStyle:
        """The style with display name `key`; ``KeyError`` when there is none.

        python-docx also accepts a style **id** here, with a deprecation
        warning; this accepts one silently, because CR-003 section 4 makes
        ``style_id`` a first-class way to name a style.
        """
        found = self._find(key)
        if found is None:
            raise KeyError(f"no style with name {key!r}")
        return style_factory(found, self._package)

    def _find(self, key: str) -> Style | None:
        wanted = key.strip().lower()
        for style in self._list:
            stored = getattr(getattr(style, "name", None), "val", None)
            style_id = style.style_id or ""
            if wanted in {
                display_name_of(style_id, stored).lower(),
                (stored or "").lower(),
                style_id.lower(),
            }:
                return style
        return None

    def default(self, style_type: WD_STYLE_TYPE) -> BaseStyle | None:
        """The default style of `style_type`, or None; the last one wins, per the spec."""
        wanted = WD_STYLE_TYPE(style_type)
        found = None
        for style in self._list:
            value = getattr(style, "type_value", None)
            value = getattr(value, "value", value)
            if value is None or str(value) != wanted.xml_value:
                continue
            if style.default:
                found = style
        return None if found is None else style_factory(found, self._package)

    def get_by_id(self, style_id: str | None, style_type: WD_STYLE_TYPE) -> BaseStyle | None:
        """The style of `style_type` with `style_id`, or the default for the type."""
        if style_id is None:
            return self.default(style_type)
        for style in self._list:
            if style.style_id == style_id:
                found = style_factory(style, self._package)
                return found if found.type is WD_STYLE_TYPE(style_type) else self.default(
                    style_type
                )
        return self.default(style_type)

    def get_style_id(self, style_or_name: Any, style_type: WD_STYLE_TYPE) -> str | None:
        """The id to write for a style object or a display name; None for None.

        None also comes back when the named style **is** the default for its
        type, which is python-docx's way of saying "write no ``w:pStyle``".
        """
        if style_or_name is None:
            return None
        if isinstance(style_or_name, BaseStyle):
            if style_or_name.type is not WD_STYLE_TYPE(style_type):
                raise ValueError(
                    f"assigned style is type {style_or_name.type}, need type "
                    f"{WD_STYLE_TYPE(style_type)}"
                )
            default = self.default(style_type)
            if default is not None and default == style_or_name:
                return None
            return style_or_name.style_id
        return self.get_style_id(self[str(style_or_name)], style_type)

    def add_style(
        self, name: str, style_type: WD_STYLE_TYPE, builtin: bool = False
    ) -> BaseStyle:
        """A new style definition named `name`, of `style_type`.

        The id is Word's: the display name with its spaces removed, and the
        built-ins' special cases (``"heading 1"`` is stored for ``Heading1``)
        taken from :data:`~docx4j_py.model.content.styles.BUILT_IN_STYLES`.
        """
        from docx4j_py.wml import el

        if name in self:
            raise ValueError(f"document already contains style {name!r}")
        styles = styles_of(self._package, create=True)
        if styles is None:
            raise ValueError("this document has no /word/styles.xml to add a style to")
        with self._recording("docx.add_style"):
            style = Style(
                style_id=style_id_of(None, name),
                type_value=WD_STYLE_TYPE(style_type).xml_value,
            )
            if not builtin:
                style.custom_style = True
            style.name = el.name(val=_stored_name(name))
            style.name.parent = style
            styles.style.append(style)
            return style_factory(style, self._package)

    @property
    def latent_styles(self) -> Any:
        """Not offered; ``styles.element.latent_styles`` is the typed ``w:latentStyles``."""
        raise AttributeError(
            "docx4j_py.docx does not offer latent_styles; the typed w:latentStyles is "
            "document.styles.element.latent_styles (CR-003 section 20)"
        )

    def _recording(self, operation: str) -> Any:
        from docx4j_py.model.content.reports import recording_on

        part = getattr(self._package, "style_definitions_part", None)
        return recording_on(self._package, operation, *([part] if part is not None else []))


def _stored_name(name: str) -> str:
    """The ``w:name`` Word stores for a display name (``"heading 1"``)."""
    from docx4j_py.model.content.styles import BUILT_IN_STYLES

    for entry in BUILT_IN_STYLES:
        if entry.name.lower() == name.strip().lower():
            return entry.stored_name or entry.name
    return name


# ---------------------------------------------------------------------------
# what the text and table proxies call
# ---------------------------------------------------------------------------


def style_for(package: Any, style_id: Any, style_type: WD_STYLE_TYPE) -> Any:
    """The style object a ``w:pStyle`` / ``w:rStyle`` / ``w:tblStyle`` names.

    The document's default for the type when the id is absent or unknown, which
    is python-docx's rule; None when the document defines no default either.
    """
    value = getattr(style_id, "val", style_id)
    return Styles(package).get_by_id(None if value is None else str(value), style_type)


def style_id_for(
    package: Any,
    style_or_name: Any,
    style_type: WD_STYLE_TYPE,
    *,
    define: bool = False,
) -> str | None:
    """The id to write for what a ``style`` setter was given.

    A style object or a display name goes through :class:`Styles`, as
    python-docx does; a plain string that names a **built-in** style the
    document does not define is resolved --- and, when `define`, defined ---
    through the content API's :func:`style_id_of`, which is CR-003 section
    14.9's rule and the reason ``add_heading(text, 7)`` works on a document
    ``create_package()`` made.
    """
    if style_or_name is None:
        return None
    if isinstance(style_or_name, BaseStyle):
        return Styles(package).get_style_id(style_or_name, style_type)
    name = str(style_or_name)
    styles = Styles(package)
    if name in styles:
        return styles.get_style_id(styles[name], style_type)
    return style_id_of(package, name, validate=True, define=define)


_ = recording
