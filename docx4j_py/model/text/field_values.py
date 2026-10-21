"""Field results computed under ``fields: update`` (WD REQ-506). CR-011 section 6.

docx4j's ``org.docx4j.model.text.FieldValues`` at ``198d015a3``: in the walk,
and so without a copy of the package, only the fields that need neither layout
nor external access, by ``FieldUpdater``'s resolvers and
``FormattingSwitchHelper``'s switches (``\\*``, ``\\#``, ``\\@``;
:mod:`docx4j_py.model.fields`):

* ``DOCPROPERTY`` (core, application and custom properties) and ``DOCVARIABLE``;
* the core-property fields ``AUTHOR``, ``TITLE``, ``SUBJECT``, ``KEYWORDS``,
  ``COMMENTS``, ``LASTSAVEDBY``, ``REVNUM``, ``CREATEDATE`` and ``SAVEDATE``
  (one given a text argument gives that text, as Word sets the property);
* ``SEQ``, counted in reading order (``\\r``, ``\\c`` and ``\\h`` honoured;
  ``\\s`` not);
* ``REF`` to a bookmark, its text in the view (with ``\\n``, ``\\r``, ``\\w``,
  ``\\p`` or ``\\f``, which need numbering or layout, it keeps its cached
  result);
* ``QUOTE``, its argument.

Nothing else (``FORMTEXT`` is the walk's, from its form field data): not
``MERGEFIELD`` (no data is supplied), ``DATE``, ``TIME``, ``PRINTDATE``, the
layout fields or the external ones, and never a field with ``w:fldLock``. A
field whose value cannot be found keeps its cached result.
"""

from __future__ import annotations

from typing import Any

__all__ = ["FieldValues"]

#: The core-property fields, by the DOCPROPERTY name of their property (KEYWORDS read directly).
_CORE: dict[str, str | None] = {
    "AUTHOR": "Author",
    "TITLE": "Title",
    "SUBJECT": "Subject",
    "COMMENTS": "Comments",
    "LASTSAVEDBY": "LastSavedBy",
    "REVNUM": "RevisionNumber",
    "CREATEDATE": "CreateTime",
    "SAVEDATE": "LastSavedTime",
    "KEYWORDS": None,
}
#: REF's switches that need the numbering model or layout.
_REF_UNSUPPORTED = frozenset({"\\n", "\\r", "\\w", "\\p", "\\f"})


def _unquote(value: str | None) -> str | None:
    if value is None:
        return None
    text = value.strip()
    if len(text) >= 2 and text.startswith('"') and text.endswith('"'):
        text = text[1:-1]
    return text


def _switch_value(name: str, parameters: list[str]) -> str | None:
    for i in range(len(parameters) - 1):
        if parameters[i].lower() == name.lower():
            return _unquote(parameters[i + 1])
    return None


class FieldValues:
    """The computed results of one reading: its sequences and its bookmarks' text."""

    __slots__ = ("_properties", "_variables", "bookmarks", "package", "sequences")

    def __init__(self, package: Any) -> None:
        """A reader over a package; its resolvers made on first use."""
        self.package = package
        self._properties: Any = None
        self._variables: Any = None
        self.sequences: dict[str, int] = {}
        #: Bookmarks' text by name (lower case): from a first reading, then as this one reads them.
        self.bookmarks: dict[str, str] = {}

    def compute(self, instr: str | None) -> str | None:
        """A field's computed result, or None to keep its cached one."""
        from docx4j_py.model.fields import FldSimpleModel, apply_formatting_switch

        if self.package is None or instr is None:
            return None
        model = FldSimpleModel(instr)
        name = model.name
        if name is None:
            return None
        parameters = model.parameters
        argument = None if not parameters or parameters[0].startswith("\\") else _unquote(parameters[0])
        try:
            if name == "DOCPROPERTY":
                value = None if argument is None else self.properties().value(argument)
            elif name == "DOCVARIABLE":
                value = None if argument is None else self.variables().value(argument)
            elif name == "SEQ":
                value = self._sequence(argument, parameters)
            elif name == "REF":
                value = self._ref(argument, parameters)
            elif name == "QUOTE":
                value = argument
            elif name in _CORE:
                if argument is not None:
                    value = argument
                elif _CORE[name] is None:
                    value = self._keywords()
                else:
                    value = self.properties().value(_CORE[name])  # type: ignore[arg-type]
            else:
                return None
        except Exception:  # noqa: BLE001 - a field whose value cannot be found keeps its cached result
            return None
        if value is None:
            return None
        try:
            return apply_formatting_switch(self.package, model, value)
        except Exception:  # noqa: BLE001 - a switch docx4j refuses leaves the value as it is
            return value

    def properties(self) -> Any:
        if self._properties is None:
            from docx4j_py.model.fields import DocPropertyResolver

            self._properties = DocPropertyResolver(self.package)
        return self._properties

    def variables(self) -> Any:
        if self._variables is None:
            from docx4j_py.model.fields import DocVariableResolver

            self._variables = DocVariableResolver(self.package)
        return self._variables

    def _keywords(self) -> str | None:
        part = getattr(self.package, "doc_props_core_part", None)
        if part is None:
            return None
        core = part.read_contents()
        return getattr(core, "keywords", None) if core is not None else None

    def _sequence(self, identifier: str | None, parameters: list[str]) -> str | None:
        if identifier is None:
            return None
        key = identifier.upper()
        n = self.sequences.get(key, 0)
        reset = _switch_value("\\r", parameters)
        if reset is not None:
            try:
                n = int(reset.strip())
            except ValueError:
                return None
        elif "\\c" not in parameters:
            n += 1
        self.sequences[key] = n
        if "\\h" in parameters:
            return ""
        return str(n)

    def _ref(self, bookmark: str | None, parameters: list[str]) -> str | None:
        if bookmark is None:
            return None
        if any(p.lower() in _REF_UNSUPPORTED for p in parameters):
            return None
        return self.bookmarks.get(bookmark.lower())
