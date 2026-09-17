"""``document.core_properties``: the Dublin Core properties of ``/docProps/core.xml``.

CR-003 section 3.6, Phase J. python-docx's ``CoreProperties`` is sixteen
read/write scalars over ``cp:coreProperties``; the typed part CR-002 holds is
the schema's, in which ``dc:creator`` is a ``SimpleLiteral`` with a wildcard
child and ``dcterms:created`` a ``W3CDTF`` carrying an ``xsi:type``, so this
module is the translation between the two. A parsed value arrives as an
``AnyElement`` or a ``DerivedElement`` and a written one as the plain text or
the ``W3CDTF``; :func:`_text_of` and :func:`_datetime_of` read either.

The part is **created when it is absent** and related from the package, as
docx4j's ``createPackage`` relates it; ``create_package()`` always writes one,
so that only comes up for a loaded document that has none.
"""

from __future__ import annotations

import datetime as dt
from typing import Any

__all__ = ["CoreProperties", "core_properties_of"]

#: python-docx's name -> the typed field, and whether it is a date.
_TEXT_FIELDS: dict[str, str] = {
    "author": "creator",
    "category": "category",
    "comments": "description",
    "content_status": "content_status",
    "identifier": "identifier",
    "keywords": "keywords",
    "language": "language",
    "last_modified_by": "last_modified_by",
    "revision": "revision",
    "subject": "subject",
    "title": "title",
    "version": "version",
}

#: The typed fields whose value is a ``dc:``/``dcterms:`` element rather than text.
_ELEMENT_FIELDS: dict[str, str] = {
    "creator": "Creator",
    "description": "Description",
    "identifier": "Identifier",
    "language": "Language",
    "subject": "Subject",
    "title": "Title",
}


def _text_of(value: Any) -> str:
    """The text of a core property, however it was parsed or built.

    ``""`` rather than None for an absent one, which is python-docx's answer.
    """
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    inner = getattr(value, "any_element", None)
    if inner is None:
        inner = getattr(value, "value", None)
    if isinstance(inner, str):
        return inner
    text = getattr(inner, "text", None)
    if isinstance(text, str):
        return text
    if inner is None:
        return ""
    return str(inner)


def _datetime_of(value: Any) -> dt.datetime | None:
    """A ``dcterms:created`` / ``cp:lastPrinted`` as a ``datetime``, or None."""
    if value is None:
        return None
    current = value
    for _ in range(4):
        stamp = getattr(current, "to_datetime", None)
        if callable(stamp):
            return stamp()
        if isinstance(current, dt.datetime):
            return current
        nxt = getattr(current, "value", None)
        if nxt is None or nxt is current:
            return None
        current = nxt
    return None


def core_properties_of(package: Any, *, create: bool = False) -> Any | None:
    """The unmarshalled ``cp:coreProperties``, created with its part if asked.

    Returns None when the package has no core part and `create` is false.
    """
    from docx4j_py.openpackaging.parts.docprops import DocPropsCorePart

    part = None
    for candidate in getattr(package, "parts", {}).values():
        if isinstance(candidate, DocPropsCorePart):
            part = candidate
            break
    if part is None:
        if not create:
            return None
        from docx4j_py.docprops.core import CoreProperties as CorePropertiesElement
        from docx4j_py.model.content.picture import note_added_part
        from docx4j_py.model.content.reports import recording_on

        with recording_on(package, "docx.core_properties.create"):
            part = DocPropsCorePart()
            part.set_contents(CorePropertiesElement())
            had_override = package.content_type_manager.get_override_content_type(part.part_name)
            rel = package.add_target_part(part)
            note_added_part(
                package, part, rel, package, added_content_type=had_override is None
            )
    return part.contents


class CoreProperties:
    """``document.core_properties``: python-docx's sixteen scalars."""

    __slots__ = ("_package",)

    def __init__(self, package: Any) -> None:
        """Wrap the package whose core part this is over."""
        self._package = package

    def __repr__(self) -> str:
        """``<docx4j_py.docx CoreProperties 'Report' by 'Ada'>``."""
        return f"<docx4j_py.docx CoreProperties {self.title!r} by {self.author!r}>"

    @property
    def element(self) -> Any:
        """The typed ``cp:coreProperties``; created with its part when absent."""
        return core_properties_of(self._package, create=True)

    def _read(self, field: str) -> Any:
        properties = core_properties_of(self._package)
        return None if properties is None else getattr(properties, field, None)

    def _text(self, name: str) -> str:
        return _text_of(self._read(_TEXT_FIELDS[name]))

    def _set_text(self, name: str, value: str) -> None:
        field = _TEXT_FIELDS[name]
        with self._recording(f"docx.core_properties.{name}"):
            properties = core_properties_of(self._package, create=True)
            if properties is None:  # pragma: no cover - create=True always answers
                return
            wrapper = _ELEMENT_FIELDS.get(field)
            if wrapper is None:
                setattr(properties, field, value)
                return
            import docx4j_py.docprops.core as core_module

            item = getattr(core_module, wrapper)(any_element=value)
            setattr(properties, field, item)
            item.parent = properties

    def _date(self, field: str) -> dt.datetime | None:
        return _datetime_of(self._read(field))

    def _set_date(self, field: str, value: dt.datetime | None, *, w3cdtf: bool) -> None:
        from docx4j_xsdata.models.datatype import XmlDateTime

        with self._recording(f"docx.core_properties.{field}"):
            properties = core_properties_of(self._package, create=True)
            if properties is None:  # pragma: no cover
                return
            if value is None:
                setattr(properties, field, None)
                return
            moment = value if value.tzinfo is not None else value.replace(tzinfo=dt.UTC)
            stamp = XmlDateTime.from_datetime(
                moment.astimezone(dt.UTC).replace(microsecond=0)
            )
            if w3cdtf:
                from docx4j_py.docprops.dcterms import W3CDTF

                item = W3CDTF(value=stamp)
                setattr(properties, field, item)
                item.parent = properties
            else:
                setattr(properties, field, stamp)

    def _recording(self, operation: str) -> Any:
        from docx4j_py.model.content.reports import recording_on
        from docx4j_py.openpackaging.parts.docprops import DocPropsCorePart

        part = next(
            (
                candidate
                for candidate in getattr(self._package, "parts", {}).values()
                if isinstance(candidate, DocPropsCorePart)
            ),
            None,
        )
        return recording_on(self._package, operation, *([part] if part is not None else []))

    # -- the text properties ----------------------------------------------

    @property
    def author(self) -> str:
        """``dc:creator``."""
        return self._text("author")

    @author.setter
    def author(self, value: str) -> None:
        self._set_text("author", value)

    @property
    def category(self) -> str:
        """``cp:category``."""
        return self._text("category")

    @category.setter
    def category(self, value: str) -> None:
        self._set_text("category", value)

    @property
    def comments(self) -> str:
        """``dc:description``, which Word shows as *Comments*."""
        return self._text("comments")

    @comments.setter
    def comments(self, value: str) -> None:
        self._set_text("comments", value)

    @property
    def content_status(self) -> str:
        """``cp:contentStatus``."""
        return self._text("content_status")

    @content_status.setter
    def content_status(self, value: str) -> None:
        self._set_text("content_status", value)

    @property
    def identifier(self) -> str:
        """``dc:identifier``."""
        return self._text("identifier")

    @identifier.setter
    def identifier(self, value: str) -> None:
        self._set_text("identifier", value)

    @property
    def keywords(self) -> str:
        """``cp:keywords``."""
        return self._text("keywords")

    @keywords.setter
    def keywords(self, value: str) -> None:
        self._set_text("keywords", value)

    @property
    def language(self) -> str:
        """``dc:language``."""
        return self._text("language")

    @language.setter
    def language(self, value: str) -> None:
        self._set_text("language", value)

    @property
    def last_modified_by(self) -> str:
        """``cp:lastModifiedBy``."""
        return self._text("last_modified_by")

    @last_modified_by.setter
    def last_modified_by(self, value: str) -> None:
        self._set_text("last_modified_by", value)

    @property
    def subject(self) -> str:
        """``dc:subject``."""
        return self._text("subject")

    @subject.setter
    def subject(self, value: str) -> None:
        self._set_text("subject", value)

    @property
    def title(self) -> str:
        """``dc:title``."""
        return self._text("title")

    @title.setter
    def title(self, value: str) -> None:
        self._set_text("title", value)

    @property
    def version(self) -> str:
        """``cp:version``."""
        return self._text("version")

    @version.setter
    def version(self, value: str) -> None:
        self._set_text("version", value)

    @property
    def revision(self) -> int:
        """``cp:revision`` as an integer; 0 when absent or not a number."""
        text = self._text("revision")
        try:
            return int(text)
        except ValueError:
            return 0

    @revision.setter
    def revision(self, value: int) -> None:
        if not isinstance(value, int) or value < 1:
            raise ValueError("revision property requires a positive int value")
        self._set_text("revision", str(value))

    # -- the dates ---------------------------------------------------------

    @property
    def created(self) -> dt.datetime | None:
        """``dcterms:created``."""
        return self._date("created")

    @created.setter
    def created(self, value: dt.datetime | None) -> None:
        self._set_date("created", value, w3cdtf=True)

    @property
    def modified(self) -> dt.datetime | None:
        """``dcterms:modified``."""
        return self._date("modified")

    @modified.setter
    def modified(self, value: dt.datetime | None) -> None:
        self._set_date("modified", value, w3cdtf=True)

    @property
    def last_printed(self) -> dt.datetime | None:
        """``cp:lastPrinted``."""
        return self._date("last_printed")

    @last_printed.setter
    def last_printed(self, value: dt.datetime | None) -> None:
        self._set_date("last_printed", value, w3cdtf=False)
