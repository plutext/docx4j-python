"""Field results computed under ``fields: update`` (WD REQ-506). CR-011 section 6; CR-025 Phase B.

docx4j's ``org.docx4j.model.text.FieldValues`` at ``c3e9920e5``: in the walk,
and so without a copy of the package, only the fields that need neither layout
nor external access, through the evaluators by name
(:mod:`docx4j_py.model.fields.evaluators`, docx4j's ``FieldEvaluators``, its
CR-067) with this walk's :class:`~docx4j_py.model.fields.FieldContext`:

* ``DOCPROPERTY`` (core, application and custom properties) and ``DOCVARIABLE``;
* the core-property fields ``AUTHOR``, ``TITLE``, ``SUBJECT``, ``KEYWORDS``,
  ``COMMENTS``, ``LASTSAVEDBY`` and ``REVNUM`` (one given a text argument gives
  that text, as Word sets the property), and ``CREATEDATE`` and ``SAVEDATE``,
  the property's instant to the minute in the package's zone
  (``pkg.date_time_zone``, else the process's), in Word's default format
  where the field has no ``\\@`` (``2/3/2024 3:05:00 PM``);
* ``SEQ``, counted in reading order (``\\r``, ``\\c`` and ``\\h`` honoured;
  ``\\s`` not);
* ``REF`` to a bookmark, its text in the view (with ``\\n``, ``\\r``, ``\\w``,
  ``\\p`` or ``\\f``, which need numbering or layout, it keeps its cached
  result);
* ``QUOTE``, its argument, empty where it has none;
* since docx4j CR-067 phase 1: ``INFO`` as the field it names (where that is
  one of these), ``SECTION`` (the main story's, counted at the walk),
  ``NOTEREF`` (the number of the note whose reference the bookmark holds),
  ``USERNAME``, ``USERINITIALS`` and ``USERADDRESS`` where the caller supplies
  the user (``TextOptions(field_user=...)``; never by default, as docx4j's
  extraction has none).

Nothing else (``FORMTEXT`` is the walk's, from its form field data): not
``MERGEFIELD`` (no data is supplied), ``DATE``, ``TIME``, ``PRINTDATE`` (a
"now" would make the text change with the clock; CR-025 D14), the layout fields
or the external ones, and never a field with ``w:fldLock``. A field whose
value cannot be found keeps its cached result. What only a write-back shows
is not here: a ``REF``'s result in the bookmark's formatting, ``NOTEREF \\f``'s
character style, ``\\h``'s hyperlink.
"""

from __future__ import annotations

from typing import Any

__all__ = ["COMPUTED", "FieldValues"]

#: The names computed here: the WD's set, which docx4j CR-067's phases widen through the WD (docx4j ``COMPUTED``).
COMPUTED: frozenset[str] = frozenset(
    {
        "DOCPROPERTY",
        "DOCVARIABLE",
        "SEQ",
        "REF",
        "QUOTE",
        "AUTHOR",
        "TITLE",
        "SUBJECT",
        "COMMENTS",
        "LASTSAVEDBY",
        "REVNUM",
        "CREATEDATE",
        "SAVEDATE",
        "KEYWORDS",
        # docx4j CR-067 phase 1
        "INFO",
        "SECTION",
        "NOTEREF",
        "USERNAME",
        "USERINITIALS",
        "USERADDRESS",
    }
)


class FieldValues:
    """The computed results of one reading: its context, with its sequences and its bookmarks' text."""

    __slots__ = ("context",)

    def __init__(self, package: Any, *, document: Any = None, user: Any = None) -> None:
        """A reader over a package; `document` is the ``w:document`` the walk reads, `user` the user fields' user."""
        from docx4j_py.model.fields.context import FieldContext

        #: This reading's context (docx4j ``FieldContext.of(pkg)``), or None with no package.
        self.context: Any = (
            None if package is None else FieldContext(package, user=user, document=document, computes=COMPUTED)
        )

    @property
    def bookmarks(self) -> dict[str, str]:
        """Bookmarks' text by name (lower case): from a first reading, then as this one reads them."""
        return self.context.bookmarks if self.context is not None else {}

    @property
    def sequences(self) -> dict[str, int]:
        """``SEQ`` counters by identifier (upper case), this reading's."""
        return self.context.sequences if self.context is not None else {}

    def compute(self, instr: str | None, section: int | None = None) -> str | None:
        """A field's computed result, or None to keep its cached one; `section` the main story's (from 1), or None."""
        from docx4j_py.model.fields import FieldInstruction, apply_formatting_switch
        from docx4j_py.model.fields import evaluators as registry

        context = self.context
        if context is None or instr is None:
            return None
        context.section = section
        instruction = FieldInstruction(instr)
        name = instruction.name
        if name is None or name not in COMPUTED:
            return None
        evaluator = registry.get(name)
        if evaluator is None:
            return None
        try:
            value = evaluator(instruction, context)
        except Exception:  # noqa: BLE001 - a field whose value cannot be found keeps its cached result
            return None
        if value is None or registry.formats_itself(name):
            return value
        try:
            return apply_formatting_switch(context.package, instruction, value)
        except Exception:  # noqa: BLE001 - a switch docx4j refuses leaves the value as it is
            return value
