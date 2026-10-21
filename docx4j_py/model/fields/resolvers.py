"""A ``DOCPROPERTY``'s and a ``DOCVARIABLE``'s value. CR-011 section 6.

docx4j's ``DocPropertyResolver`` and ``DocVariableResolver`` at
``198d015a3``, read without changing the package (each part with
``read_contents``). A name with no value raises :class:`FieldValueError`, and
the field keeps its cached result.
"""

from __future__ import annotations

from typing import Any

from docx4j_py.model.fields.formatting import FieldFormattingError, FieldValueError

__all__ = ["DocPropertyResolver", "DocVariableResolver"]

#: The core properties' DOCPROPERTY names (matched whole, so a custom property named Page or Time is still found).
_CORE = {
    "Author": "creator",
    "Comments": "description",
    "CreateTime": "created",
    "LastPrinted": "last_printed",
    "LastSavedBy": "last_modified_by",
    "LastSavedTime": "modified",
    "RevisionNumber": "revision",
    "Subject": "subject",
    "Title": "title",
}
#: The application properties' DOCPROPERTY names.
_EXTENDED = {
    "Characters": "characters",
    "CharactersWithSpaces": "characters_with_spaces",
    "Company": "company",
    "HyperlinkBase": "hyperlink_base",
    "Lines": "lines",
    "NameofApplication": "application",
    "Pages": "pages",
    "Paragraphs": "paragraphs",
    "Security": "doc_security",
    "Template": "template",
    "TotalEditingTime": "total_time",
    "Words": "words",
}
_UNSUPPORTED = frozenset({"Bytes"})


def _contents(part: Any) -> Any:
    if part is None:
        return None
    try:
        return part.read_contents()
    except Exception:  # noqa: BLE001 - a part that will not read has no values
        return None


def _first(value: Any) -> Any:
    """A Dublin Core value's first piece of text, a date's lexical form, or the value."""
    from docx4j_py.model.text.metadata import _literal

    return _literal(value)


class DocPropertyResolver:
    """The values ``DOCPROPERTY`` reads: custom, core and application properties. docx4j ``DocPropertyResolver``."""

    __slots__ = ("core", "custom", "extended")

    def __init__(self, package: Any) -> None:
        """A resolver over a package's three properties parts."""
        self.core = _contents(getattr(package, "doc_props_core_part", None))
        self.extended = _contents(getattr(package, "doc_props_extended_part", None))
        self.custom = _contents(getattr(package, "doc_props_custom_part", None))

    def value(self, key: str) -> str | None:
        """A property's value as text; None for one docx4j does not read (``Bytes``).

        Raises:
            FieldValueError: no value.
            FieldFormattingError: a custom property of a type docx4j does not convert.
        """
        value: Any = None
        if self.custom is not None:
            value = next((p for p in self.custom.property or () if p.name == key), None)
        if key in _CORE:
            value = _first(getattr(self.core, _CORE[key], None)) if self.core is not None else None
        elif key in _EXTENDED:
            value = getattr(self.extended, _EXTENDED[key], None) if self.extended is not None else None
        elif key in _UNSUPPORTED:
            return None
        if value is None:
            raise FieldValueError(f"No value found for DOCPROPERTY {key}")
        if isinstance(value, bool):
            return "true" if value else "false"
        if isinstance(value, (str, int)):
            return str(value)
        if type(value).__name__ == "PropertiesProperty":
            content = value.content
            kind = type(content).__name__
            if kind == "Lpwstr":
                return content.value
            if kind in ("Filetime", "Date"):
                from docx4j_py.model.text.metadata import date_text

                return date_text(content.value)
            raise FieldFormattingError(f"{key}: a custom property of type {kind}")
        if hasattr(value, "to_string"):
            return value.to_string()
        raise FieldFormattingError(f"{key}: no conversion of {type(value).__name__} to text")


class DocVariableResolver:
    """The settings' ``w:docVars``. docx4j ``DocVariableResolver``."""

    __slots__ = ("variables",)

    def __init__(self, package: Any) -> None:
        """A resolver over a package's settings."""
        main = getattr(package, "main_document_part", None)
        settings = _contents(getattr(main, "document_settings_part", None) if main is not None else None)
        doc_vars = getattr(settings, "doc_vars", None) if settings is not None else None
        self.variables = {v.name: v.val for v in getattr(doc_vars, "doc_var", None) or () if v.name is not None}

    def value(self, key: str) -> str:
        """A variable's value.

        Raises:
            FieldValueError: no such variable.
        """
        value = self.variables.get(key)
        if value is None:
            raise FieldValueError(f"No value found for DOCVARIABLE {key}")
        return value
