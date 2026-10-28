"""A field instruction, parsed once. CR-025 Phase B, D6.

docx4j's ``org.docx4j.model.fields.FieldInstruction`` at ``c3e9920e5`` (its
CR-067 D2, where it replaced ``FldSimpleModel``): the field's name, its
parameters as Word tokenises them (a quoted parameter keeps its quotes, so a
switch's value is told from a switch), and accessors for the argument and the
switches. The grammar is ECMA-376 17.16.2: optional white space, the field
name (or ``=`` for a formula, which needs no space after it), then arguments
and switches (``\\x``, some with a value) separated by white space, quotation
marks grouping a parameter that holds spaces. Word may split an instruction
across runs; callers join the runs before parsing.

The tokeniser is the one :class:`~docx4j_py.model.fields.FldSimpleModel`
always had (docx4j moved it into ``FieldInstruction``); ``FldSimpleModel`` is
now a subclass of this, keeping its one other name, ``parameter_string``.
"""

from __future__ import annotations

__all__ = ["MERGEFORMAT", "FieldInstruction", "split_parameters", "unquote"]

#: ``\\* MERGEFORMAT``, which is no format: a switch value skipped for a later ``\\*``.
MERGEFORMAT = "MERGEFORMAT"


def split_parameters(text: str | None) -> list[str]:
    """Word's tokenisation of the text after a field's name. docx4j ``FieldInstruction.splitParameters``.

    Parameters separated by spaces, a quoted one kept whole with its quotes.
    """
    out: list[str] = []
    if not text:
        return out
    start = -1
    quoted = False
    for i, c in enumerate(text):
        if c == '"':
            if quoted:
                out.append(text[start : i + 1])
                quoted = False
                start = -1
            else:
                start = i
                quoted = True
        elif c == " ":
            if start > -1 and not quoted:
                out.append(text[start:i])
                start = -1
        elif start == -1:
            start = i
    if start > -1:
        out.append(text[start:])
    return out


def unquote(parameter: str | None) -> str | None:
    """The parameter without its enclosing quotation marks, where it has them. docx4j ``unquote``."""
    if parameter is not None and len(parameter) > 1 and parameter[0] == '"' and parameter[-1] == '"':
        return parameter[1:-1]
    return parameter


class FieldInstruction:
    """A field instruction: its name (upper case; ``=`` for a formula), the text after it, and its parameters.

    :attr:`parameters` is a copy, so a caller cannot change what the
    instruction holds, as docx4j's unmodifiable list ensures; the other
    attributes are plain, as ``FldSimpleModel``'s always were. An empty or
    blank instruction has a None :attr:`name`.
    """

    __slots__ = ("_parameters", "name", "parameter_text", "text")

    def __init__(self, instr: str | None = None) -> None:
        """Parse an instruction, as ``w:instr`` or the joined ``w:instrText`` runs give it."""
        #: The instruction as given.
        self.text: str | None = instr
        #: The field's name, upper case (``DATE``, ``MERGEFIELD``; ``=`` for a formula), or None.
        self.name: str | None = None
        #: The text after the name, trimmed, or None.
        self.parameter_text: str | None = None
        text = instr or ""
        start = 0
        while start < len(text) and text[start] == " ":
            start += 1
        if start < len(text):
            end = start + 1
            if text[start] != "=":  # a formula is "=" alone, whatever follows it
                while end < len(text) and text[end] != " ":
                    end += 1
            self.name = text[start:end].upper()
            if end < len(text):
                self.parameter_text = text[end:].strip()
        self._parameters: tuple[str, ...] = tuple(split_parameters(self.parameter_text))

    @classmethod
    def parse(cls, instr: str | None) -> FieldInstruction:
        """docx4j's ``FieldInstruction.parse``: the same as the constructor."""
        return cls(instr)

    @property
    def parameters(self) -> list[str]:
        """The parameters as tokenised, a quoted one with its quotes; a copy, never None."""
        return list(self._parameters)

    @property
    def argument(self) -> str | None:
        """The first parameter where it is not a switch, its quotes removed; None where there is none.

        ``DATE \\@ "d MMMM"`` has none; ``MERGEFIELD "First Name"`` has ``First Name``.
        """
        if not self._parameters:
            return None
        first = self._parameters[0]
        if first.startswith("\\"):
            return None
        return unquote(first)

    def has_switch(self, switch: str) -> bool:
        """Whether the switch (``"\\h"``) is present; the match is exact, as Word's is for the letter switches."""
        return self.index_of_switch(switch, 0) > -1

    def index_of_switch(self, switch: str, start: int = 0) -> int:
        """The index of the switch's first occurrence at or after `start`, or -1."""
        for i in range(max(0, start), len(self._parameters)):
            if self._parameters[i] == switch:
                return i
        return -1

    def value_at(self, index: int) -> str | None:
        """The parameter at `index`, its quotes removed; None where there is none or it is a switch."""
        if index < 0 or index >= len(self._parameters):
            return None
        value = unquote(self._parameters[index]) or ""
        if value.startswith("\\"):
            return None
        return value

    def switch_value(self, switch: str) -> str | None:
        """The value of the switch's first occurrence: None where it is absent, ``""`` where it has none.

        A ``\\* MERGEFORMAT`` is skipped in favour of a later ``\\*`` with a
        format, as the switches read the general switch.
        """
        position = self.index_of_switch(switch, 0)
        if position < 0:
            return None
        value = self.value_at(position + 1)
        if value == MERGEFORMAT:
            value = None
            position = self.index_of_switch(switch, position + 1)
            if position > -1:
                value = self.value_at(position + 1)
        return "" if value is None else value

    def switch_values(self, switch: str) -> list[str | None]:
        """Every value of the switch, in order (None where an occurrence has none); empty where it is absent."""
        out: list[str | None] = []
        position = 0
        while (position := self.index_of_switch(switch, position)) > -1:
            position += 1
            out.append(self.value_at(position))
        return out

    def __repr__(self) -> str:
        """``FieldInstruction[DATE ['\\\\@', '"d MMMM"']]``, as docx4j's ``toString``."""
        return f"{type(self).__name__}[{self.name} {list(self._parameters)}]"
